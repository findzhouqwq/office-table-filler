@echo off
setlocal
rem 用 signtool 对安装程序 / PyInstaller 产物做代码签名。
rem 需要安装 Windows SDK（含 signtool.exe），并具备代码签名证书。
rem
rem 用法：
rem   set SIGN_PFX=D:\certs\my.pfx
rem   set SIGN_PASSWORD=******
rem   installer\sign-setup.bat installer\TableFiller-Setup.exe
rem
rem 或使用证书指纹（证书已导入当前用户证书存储）：
rem   set SIGN_THUMBPRINT=ABCDEF...
rem   installer\sign-setup.bat installer\TableFiller-Setup.exe

set "TARGET=%~1"
if "%TARGET%"=="" set "TARGET=%~dp0TableFiller-Setup.exe"
if not exist "%TARGET%" (
    echo [error] 找不到待签名文件：%TARGET%
    exit /b 1
)

where signtool >nul 2>nul
if errorlevel 1 (
    echo [error] 未找到 signtool，请安装 Windows SDK 并加入 PATH。
    exit /b 1
)

if not "%SIGN_PFX%"=="" (
    signtool sign /fd SHA256 /f "%SIGN_PFX%" /p "%SIGN_PASSWORD%" ^
        /tr http://timestamp.digicert.com /td SHA256 "%TARGET%"
) else if not "%SIGN_THUMBPRINT%"=="" (
    signtool sign /fd SHA256 /sha1 "%SIGN_THUMBPRINT%" ^
        /tr http://timestamp.digicert.com /td SHA256 "%TARGET%"
) else (
    echo [error] 请设置 SIGN_PFX + SIGN_PASSWORD，或 SIGN_THUMBPRINT。
    echo        未签名会导致 SmartScreen 和部分杀软拦截。
    exit /b 1
)

if errorlevel 1 (
    echo [error] 签名失败。
    exit /b 1
)

signtool verify /pa /v "%TARGET%"
if errorlevel 1 (
    echo [error] 签名校验失败。
    exit /b 1
)

echo [done] 已签名并校验：%TARGET%
exit /b 0
