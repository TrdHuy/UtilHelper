using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.PowerPoint;

namespace ppt2img
{
    class Program
    {
        static int width = 0;
        static int height = 0;
        static String imgType = "jpg";
        static String outDir = ".";
        static String projectRefPath = "";
        static String projectRefUrl = "";
        static String projectRefRemoteBranch = "";
        static String currentDir = Path.GetFullPath(".");
        static String currentExecuteDir = AppDomain.CurrentDomain.BaseDirectory;
        static void Main(string[] args)
        {
            Console.WriteLine($"currentExeDir = {currentExecuteDir}");
            if (args.Length == 0)
            {
                Console.WriteLine(@"Usage: ppt2img <ppt|pptx> [options]
            Option:
                -rB|--ref-remote-branch (required)
                -rU|--ref-repo-url (required)
                -rP|--ref-repo-path (required)
                -t|--type <png|jpg>
            Example:
                ppt2img ""C:\SCS\VSProj\SPRNetTool\trdhuy\ArtWiz\UIUX\[feature]_png_to_spr.pptx"" -rB ""origin/main"" -rU ""https://github.sec.samsung.net/huy-td1/test_repo"" -rP ""C:\SCS\VSProj\SPRNetTool\trdhuy\test_repo""");
                Environment.Exit(1);
                return;
            }

            var listInputFile = new List<string>();
            try
            {
                for (int i = 0; i < args.Length; ++i)
                {
                    if (args[i] == "--type" || args[i] == "-t")
                    {
                        ++i;
                        imgType = args[i];
                    }
                    else if (args[i] == "--ref-repo-path" || args[i] == "-rP")
                    {
                        ++i;
                        projectRefPath = FormatPath(args[i]);
                    }
                    else if (args[i] == "--ref-repo-url" || args[i] == "-rU")
                    {
                        ++i;
                        projectRefUrl = args[i];
                    }
                    else if (args[i] == "--ref-remote-branch" || args[i] == "-rB")
                    {
                        ++i;
                        projectRefRemoteBranch = args[i];
                    }
                    else
                    {
                        listInputFile.Add(FormatPath(args[i]));
                    }
                }
                Console.WriteLine("listInputFile = ");
                foreach (var inpt in listInputFile)
                {
                    Console.WriteLine(inpt);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Invalid args");
                Console.WriteLine("{0}", e.Message);
                Environment.Exit(1);
                return;
            }
            if (listInputFile.Count == 0)
            {
                Console.WriteLine("Missing input file!");
                Environment.Exit(1);
                return;
            }
            if (projectRefUrl == "")
            {
                Console.WriteLine("Missing -rU required option!");
                Environment.Exit(1);
                return;
            }
            if (projectRefRemoteBranch == "" && projectRefPath != "")
            {
                Console.WriteLine("Missing -rB required option with -rP option!");
                Environment.Exit(1);
                return;
            }

            outDir = Path.GetFullPath(".") + "\\.mdcache";

            if (Directory.Exists(outDir))
            {
                Directory.Delete(outDir, true);
            }
            if (!Directory.Exists(outDir))
            {
                Directory.CreateDirectory(outDir);
            }
            var cacheGitFilePath = $"{currentExecuteDir}\\.cache";
            if (!Directory.Exists(cacheGitFilePath))
            {
                Directory.CreateDirectory(cacheGitFilePath);
            }

            try
            {
                List<(int, List<string>, string)> slideCountAndOutNames = new List<(int, List<string>, string)>();
                foreach (var inPpt in listInputFile)
                {
                    int slideCount;
                    List<string> outNames = new List<string>();
                    ExportPpt2Img(inPpt, out slideCount, outNames);
                    slideCountAndOutNames.Add((slideCount, outNames, inPpt));
                }

                string pushRefCmd = buildPushRefCmd(projectRefPath, projectRefRemoteBranch, $"img{DateTime.Now.Ticks}", outDir);
                ExeCmd(pushRefCmd, out string output, out string error);

                var hashid = GetHashIdFromOutput(output);
                Console.WriteLine($"hashId = {hashid}");
                if (hashid == "")
                {
                    Console.WriteLine("Empty hash id");
                    Environment.Exit(1);
                    return;
                }
                PushRefAndUpdateMdFile(slideCountAndOutNames, hashid);

                Console.WriteLine("Done");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                Environment.Exit(1);
            }
        }

        private static string FormatPath(string rawPath)
        {
            rawPath = rawPath.Replace("\"","");
            if (rawPath.StartsWith("/") || rawPath.Contains("/"))
            {
                if (rawPath.StartsWith("/c/"))
                {
                    rawPath = rawPath.Replace("/c/", "c:/");
                    rawPath = rawPath.Replace('/', '\\');
                    return rawPath;
                }
            }
            return rawPath;
        }

        private static void PushRefAndUpdateMdFile(List<(int, List<string>, string)> slideCountAndOutNames, string hashid)
        {
            foreach (var item in slideCountAndOutNames)
            {
                var inPpt = item.Item3;
                var baseName = Path.GetFileNameWithoutExtension(inPpt);
                var mdFilePath = Path.GetDirectoryName(inPpt) + "\\" + baseName + ".md";

                using (StreamWriter writer = new StreamWriter(mdFilePath))
                {
                    var currentSection = -1;

                    for (int i = 0; i < item.Item1; i++)
                    {
                        var outName = item.Item2[i];
                        var refFileUrl = Uri.EscapeUriString(GetRefFileUrl(projectRefUrl, hashid, outName));
                        var refFileName = Path.GetFileName(outName);

                        var fileNameRegex = @"(.*)_-_-_(.*)_-_-_(\d+)_-_-_sec(\d*)slide(\d+).(jpg|png|jpeg)";
                        var match = Regex.Match(refFileName, fileNameRegex);
                        if (match.Success)
                        {
                            var secName = match.Groups[2].Value;
                            var matchedFileId = match.Groups[3].Value;
                            var secIndex = Convert.ToInt32(match.Groups[4].Value);
                            var slideIndex = Convert.ToInt32(match.Groups[5].Value);

                            if (secIndex != currentSection)
                            {
                                if (i > 0)
                                {
                                    writer.WriteLine("");
                                }
                                writer.WriteLine($"## {secName}\n");
                                currentSection = secIndex;
                            }
                            writer.WriteLine($"![{matchedFileId}_{secIndex}_{slideIndex}]({refFileUrl})");
                        }
                    }
                }
            }
        }

        private static void ExportPpt2Img(string inPpt, out int slideCount, List<string> outNames)
        {
            Console.WriteLine($"inPPt={inPpt}");
            var baseName = Path.GetFileNameWithoutExtension(inPpt);
            Application PowerPoint_App = new Application();
            Presentations multi_presentations = PowerPoint_App.Presentations;
            Presentation presentation = multi_presentations.Open(inPpt
                , MsoTriState.msoTrue /* ReadOnly=true */
                , MsoTriState.msoTrue /* Untitled=true */
                , MsoTriState.msoFalse /* WithWindow=false */);

            slideCount = 0;
            var fileId = DateTime.Now.Ticks;
            var hasSection = presentation.SectionProperties.Count > 0;
            slideCount = presentation.Slides.Count;
            for (int i = 0; i < slideCount; i++)
            {
                var slide = presentation.Slides[i + 1];
                var sectionName = hasSection ? presentation.SectionProperties.Name(slide.sectionIndex) : "";

                String outName = String.Format($"{outDir}\\{baseName}_-_-_{sectionName}_-_-_{fileId}_-_-_sec{slide.sectionIndex}slide{i}.{imgType}");
                Console.WriteLine("Saving slide {0} of {1}... to: {2}", i + 1, slideCount, outName);

                outNames.Add(outName);
                try
                {
                    slide.Export(outName, imgType, width, height);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Failed to export slide {0}", i + 1);
                    Console.WriteLine("{0}", e.Message);
                    break;
                }
            }

            PowerPoint_App.Quit();
        }

        static string buildPushRefCmd(string projectRefPath, string projectRefRemoteBranch, string newRandomBranch, string folderPathToPush)
        {
            if (projectRefPath == "")
            {
                var cacheGitFilePath = $"{currentExecuteDir}\\.cache";

                return $"echo \"Start create new ref\"" +
                    $"&& cd \"{cacheGitFilePath}\" " +
                    $"&& git clone \"{projectRefUrl.Trim()}.git\" " +
                    $"&& cd \"{projectRefUrl.Substring(projectRefUrl.LastIndexOf('/') + 1).Trim()}\" " +
                    $"&& git checkout --orphan {newRandomBranch} " +
                    $"&& xcopy /E /I /Y \"{folderPathToPush}\" \"changes\\{Path.GetFileName(folderPathToPush)}\" " +
                    "&& git add changes " +
                    "&& git commit -m \"add changes\" " +
                    $"&& git push origin HEAD:refs/{newRandomBranch}/image-ref " +
                    $"&& echo \"=========THIS IS COMMIT HASH ID=========\" " +
                    $"&& git log --format=%H " +
                    $"&& echo \"=========END COMMIT HASH ID=========\" " +
                    $"&& cd .. && rmdir /s /q \"{projectRefUrl.Substring(projectRefUrl.LastIndexOf('/') + 1).Trim()}\"";
            }
            return $"echo \"Start create new ref\"" +
                $"&& cd \"{projectRefPath}\" " +
                $"&& git checkout {projectRefRemoteBranch} --orphan {newRandomBranch} " +
                $"&& xcopy /E /I /Y \"{folderPathToPush}\" \"changes\\{Path.GetFileName(folderPathToPush)}\" " +
                "&& git add changes " +
                "&& git commit -m \"add changes\" " +
                $"&& git push origin HEAD:refs/{newRandomBranch}/image-ref " +
                $"&& echo \"=========THIS IS COMMIT HASH ID=========\" " +
                $"&& git log --format=%H " +
                $"&& echo \"=========END COMMIT HASH ID=========\" ";
        }

        private static string GetRefFileUrl(string gitRepoUrl, string hashId, string localFileName)
        {
            return $"{gitRepoUrl}/blob/{hashId}/changes/{Path.GetFileName(Path.GetDirectoryName(localFileName))}/{Path.GetFileName(localFileName)}?raw=true";
        }

        private static string GetHashIdFromOutput(string output)
        {
            var regex = @"""=========THIS IS COMMIT HASH ID=========""\s*\r*\n([a-zA-Z0-9]+)\r*\n""=========END COMMIT HASH ID=========""";
            var matches = Regex.Match(output, regex);
            if (matches.Success)
            {
                return matches.Groups[1].Value;
            }
            return "";
        }

        static void ExeCmd(string command, out string Output, out string Error)
        {
            string output = "";
            string error = "";
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo()
                {
                    FileName = "cmd.exe",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    RedirectStandardInput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                };
                using (Process p = new Process { StartInfo = psi })
                {

                    p.OutputDataReceived += (s, e) =>
                    {
                        output += e.Data + "\n";
                        Console.WriteLine(e.Data);
                    };

                    p.ErrorDataReceived += (s, e) =>
                    {
                        error += e.Data + "\n";
                        Console.WriteLine(e.Data);
                    };
                    p.Start();
                    p.BeginOutputReadLine();
                    p.BeginErrorReadLine();
                    using (StreamWriter sw = p.StandardInput)
                    {
                        if (sw.BaseStream.CanWrite)
                        {
                            sw.WriteLine(command);
                        }
                    }
                    p.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed message={ex.Message}");
            }
            Output = output;
            Error = error;
        }
    }
}
