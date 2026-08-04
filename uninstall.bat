@echo off
setlocal

set SOLUTION_DIR=%~dp0
set DLL_PATH=%SOLUTION_DIR%Fuwoa.AddIn\bin\Release\Fuwoa.AddIn.dll

echo ============================================
echo  FUWOA Excel Add-In Uninstall Script
echo ============================================
echo.

if not exist "%DLL_PATH%" (
    echo ERROR: DLL not found at %DLL_PATH%
    echo Please build the project first, or manually unregister from registry.
    pause
    exit /b 1
)

echo Unregistering COM component...
C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe "%DLL_PATH%" /unregister

if %ERRORLEVEL% NEQ 0 (
    echo Unregistration failed. Please run as Administrator.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo Uninstall complete. Restart Excel.
pause
