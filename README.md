# UtilHelper

### Download latest ppt2img

👉 Download & install in Downloads folder
```powershell
powershell -Command "$url = (Invoke-RestMethod -Uri 'https://api.github.com/repos/TrdHuy/UtilHelper/releases/latest' | Select-Object -ExpandProperty assets | Where-Object { $_.name -eq 'Release.zip' }).browser_download_url; Invoke-WebRequest -Uri $url -OutFile \"$env:USERPROFILE\Downloads\Release.zip\"; Expand-Archive -Path \"$env:USERPROFILE\Downloads\Release.zip\" -DestinationPath \"$env:USERPROFILE\Downloads\ppt2img\" -Force; Remove-Item -Path \"$env:USERPROFILE\Downloads\Release.zip\" -Force; [Environment]::SetEnvironmentVariable( \"Path\", \"$env:USERPROFILE\Downloads\ppt2img\Release\", \"User\")"
```

👉 Download and install according to the path of your choice
```powershell
powershell -Command "$downloadPath = Read-Host 'Nhập đường dẫn tải về:'; if (Test-Path -PathType Container -Path $downloadPath) { $url = (Invoke-RestMethod -Uri 'https://api.github.com/repos/TrdHuy/UtilHelper/releases/latest' | Select-Object -ExpandProperty assets | Where-Object { $_.name -eq 'Release.zip' }).browser_download_url; Invoke-WebRequest -Uri $url -OutFile \"$downloadPath\Release.zip\"; Expand-Archive -Path \"$downloadPath\Release.zip\" -DestinationPath \"$downloadPath\ppt2img\" -Force; Remove-Item -Path \"$downloadPath\Release.zip\" -Force; [Environment]::SetEnvironmentVariable( \"Path\", \"$downloadPath\ppt2img\Release\", \"User\" ); } else { Write-Host 'Thư mục không tồn tại.'}" 
```
