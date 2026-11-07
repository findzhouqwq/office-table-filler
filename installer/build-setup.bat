@echo off
setlocal
cd /d "%~dp0"

echo [build] 检查 ps2exe ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "if (-not (Get-Module -ListAvailable -Name ps2exe)) { Install-Module ps2exe -Scope CurrentUser -Force -AllowClobber -Repository PSGallery }"
if errorlevel 1 goto :err

echo [build] 编译无控制台安装程序 ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-ps2exe -inputFile '%~dp0TableFillerSetup.ps1' -outputFile '%~dp0TableFiller-Setup.exe' -noConsole -sta -iconFile '%~dp0TableFiller.ico' -title '表格批量填写工具 安装程序' -product '表格批量填写工具' -company 'findzhouqwq' -version '1.1.0.0'"
if errorlevel 1 goto :err

echo.
echo [done] 已生成 installer\TableFiller-Setup.exe
exit /b 0

:err
echo.
echo [error] 编译失败。请确认已安装 PowerShell 5.1+ 并能访问 PowerShell Gallery。
exit /b 1
