@echo off
echo Registering FUWOA Excel Add-In...
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe "%~dp0Fuwoa.AddIn.dll" /codebase
if %ERRORLEVEL% EQU 0 (
    echo Registration successful.
) else (
    echo Registration failed. Please run as Administrator.
)
pause
