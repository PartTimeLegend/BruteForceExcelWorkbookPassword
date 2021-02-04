FROM mcr.microsoft.com/windows/servercore:1809
LABEL maintainer="hi@antonybailey.net"
RUN ["powershell", "New-Item", "-Path \"C:\"", "-ItemType \"directory\"", "-Name \"temp\""]
WORKDIR C:/temp
COPY BruteForceExcelWorkbookPassword.ps1 c:/temp/
COPY book1.xlsx c:/temp/
RUN powershell.exe -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command "[System.Net.ServicePointManager]::SecurityProtocol = 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"
RUN choco install office2019proplus -y
RUN powershell.exe -ExecutionPolicy Bypass c:\temp\BruteForceExcelWorkbookPassword.ps1