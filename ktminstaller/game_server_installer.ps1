param (
     [string]$INSTALL_PATH,
     [string]$GITHUB_TOKEN
)

function Get-LatestReleaseInfo {
     param (
          [string]$Owner,
          [string]$Repo,
          [string]$Token
     )
     try {
          $response = Invoke-RestMethod -Uri "https://api.github.com/repos/$Owner/$Repo/releases/latest" -Headers @{ Authorization = "token $Token" }
     }
     catch {
          return $null
     }
     return $response
}
function Get-AssetIdFromReleaseNote {
     param (
          [string]$Repo,
          [string]$releaseNote
     )
     $regex = "(\|__" + $Repo + "__\|.+\|)(\d+)\|"
     $match = $releaseNote -match $regex

     if ($match) {
          return $matches[2]
     }
     else {
          return $null
     }
}
function Download-LatestReleaseInfo {
     param (
          [string]$Token,
          [string]$Repo,
          [string]$Owner,
          [string]$DownloadUrl,
          [string]$Id,
          [string]$DownloadPath,
          [string]$ExtractPath
          
     )
     $headers = @{"Authorization" = "token $Token"; "Accept" = "application/octet-stream" }
     $download = "https://api.github.com/repos/$Owner/$Repo/releases/assets/$Id"
     $response = Invoke-WebRequest -Uri $download -OutFile $DownloadPath -Headers $headers
     Expand-Archive -Path $DownloadPath -DestinationPath $ExtractPath -Force
     Remove-Item -Path $DownloadPath
     # Xóa các file trùng tên tại đích trước khi di chuyển
     $buildItems = Get-ChildItem -Path "$ExtractPath\Build" -Recurse | ForEach-Object { $_.FullName.Replace('Build\', '') }
     Get-ChildItem -Path $ExtractPath -Recurse | Where-Object {
          # Kiểm tra xem tệp hoặc thư mục có nằm trong danh sách cần xóa không
          $itemToRemove = $_.FullName
          $shouldBeRemoved = $itemToRemove -in $buildItems
          Write-Host "Checking item: $itemToRemove"
          if ($shouldBeRemoved) {
               # Hiển thị tên tệp sẽ bị xóa
               Write-Host "|__Removing item: $itemToRemove"
          }
      
          return $shouldBeRemoved
     } | Remove-Item -Force -Recurse

     Move-Item -Path "$ExtractPath\Build\*" -Destination "$ExtractPath" -Force
     Remove-Item -Path "$ExtractPath\Build"
     return $response
}

function Download-LatestReleaseInfoNoRescure {
     param (
          [string]$Token,
          [string]$Repo,
          [string]$Owner,
          [string]$DownloadUrl,
          [string]$Id,
          [string]$DownloadPath,
          [string]$ExtractPath
          
     )
     $headers = @{"Authorization" = "token $Token"; "Accept" = "application/octet-stream" }
     $download = "https://api.github.com/repos/$Owner/$Repo/releases/assets/$Id"
     $response = Invoke-WebRequest -Uri $download -OutFile $DownloadPath -Headers $headers
     Expand-Archive -Path $DownloadPath -DestinationPath $ExtractPath -Force
     Remove-Item -Path $DownloadPath

     # Xóa các file trùng tên tại đích trước khi di chuyển
     $buildItems = Get-ChildItem -Path "$ExtractPath\Build" | ForEach-Object { $_.FullName.Replace('Build\', '') }
     Get-ChildItem -Path $ExtractPath | Where-Object {
          # Kiểm tra xem tệp hoặc thư mục có nằm trong danh sách cần xóa không
          $itemToRemove = $_.FullName
          $shouldBeRemoved = $itemToRemove -in $buildItems
          Write-Host "Checking item: $itemToRemove"
          if ($shouldBeRemoved) {
               # Hiển thị tên tệp sẽ bị xóa
               Write-Host "|__Removing item: $itemToRemove"
          }
      
          return $shouldBeRemoved
     } | Remove-Item -Force -Recurse

     Move-Item -Path "$ExtractPath\Build\*" -Destination "$ExtractPath" -Force
     Remove-Item -Path "$ExtractPath\Build"
     return $response
}

function Download-AndInstall() {
     param (
          $repo,
          $assetId,
          $installPath,
          $isResecure
     )
     if (-not $assetId) {
          throw 'Failed to install GameServer'
     }
     if (-Not (Test-Path -Path $installPath)) {
          New-Item -Path $installPath -ItemType Directory
     }
     Write-Host "Download and Installing $repo at: $installPath"

     if (!$isResecure) {
          Download-LatestReleaseInfoNoRescure -Token $GITHUB_TOKEN `
               -Owner TrdHuy `
               -Repo $repo `
               -DownloadPath "$installPath/t.zip" `
               -ExtractPath "$installPath" `
               -Id $assetId
     }
     else {
          Download-LatestReleaseInfo -Token $GITHUB_TOKEN `
               -Owner TrdHuy `
               -Repo $repo `
               -DownloadPath "$installPath/t.zip" `
               -ExtractPath "$installPath" `
               -Id $assetId
     }
     
     Write-Host "Install $repo successfully at: $installPath"
}

if (-not $INSTALL_PATH) {
     Write-Host 'To run this in cmd: powershell -File SetUpServer.ps1 -INSTAL_PATH <path/to/your/sv> -GITHUB_TOKEN <your github token>'
     throw 'missing install path'
}
if (-not $GITHUB_TOKEN) {
     Write-Host 'To run this in cmd: powershell -File SetUpServer.ps1 -INSTAL_PATH <path/to/your/sv> -GITHUB_TOKEN <your github token>'
     throw 'missing GITHUB_TOKEN'
}

$GAME_SERVER_PATH = "$INSTALL_PATH\Server\GameServer\GameServer\bin\x64\Release"
$GAME_SERVER_EXTENSION_PATH = "$INSTALL_PATH\Server\GameServer\GameServer\bin\x64\Release"
$GAME_DBSERVER_PATH = "$INSTALL_PATH\Server\GameDBServer\GameDBServer\bin\Release"
$LOG_DB_SERVER_PATH = "$INSTALL_PATH\Server\LogDBServer\bin\Release"

Write-Host GAME_SERVER_PATH: $GAME_SERVER_PATH
Write-Host GAME_SERVER_EXTENSION_PATH: $GAME_SERVER_EXTENSION_PATH
Write-Host GAME_DBSERVER_PATH: $GAME_DBSERVER_PATH
Write-Host LOG_DB_SERVER_PATH: $LOG_DB_SERVER_PATH

$response = Get-LatestReleaseInfo -Owner TrdHuy -Repo KiemTheMobile -Token $GITHUB_TOKEN
$mReleaseNote = $response.body
if (-not $response) {
     throw "Failed to get realease note from KiemTheMobile"
}

#######################################################
$assetId = Get-AssetIdFromReleaseNote -Repo "KiemTheMobile.GameServer" -releaseNote $mReleaseNote
Download-AndInstall -repo "KiemTheMobile.GameServer" `
     -assetId $assetId `
     -installPath $GAME_SERVER_PATH `
     -isResecure $true
#######################################################
$assetId = Get-AssetIdFromReleaseNote -Repo "KiemTheMobile.GameServer.Extension" -releaseNote $mReleaseNote
Download-AndInstall -repo "KiemTheMobile.GameServer.Extension" `
     -assetId $assetId `
     -installPath $GAME_SERVER_PATH `
     -isResecure $false
#######################################################
$assetId = Get-AssetIdFromReleaseNote -Repo "KiemTheMobile.GameDBServer" -releaseNote $mReleaseNote
Download-AndInstall -repo "KiemTheMobile.GameDBServer" `
     -assetId $assetId `
     -installPath $GAME_DBSERVER_PATH `
     -isResecure $true
#######################################################
$assetId = Get-AssetIdFromReleaseNote -Repo "KiemTheMobile.LogDBServer" -releaseNote $mReleaseNote
Download-AndInstall -repo "KiemTheMobile.LogDBServer" `
     -assetId $assetId `
     -installPath $LOG_DB_SERVER_PATH `
     -isResecure $true
