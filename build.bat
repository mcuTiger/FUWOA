@echo off
setlocal enabledelayedexpansion

set SOLUTION_DIR=%~dp0
set CONFIG=Release

echo ============================================
echo  FUWOA Excel Add-In Build Script
echo ============================================
echo.

REM Step 1: Generate strong name key
echo [1/4] Generating strong name key...
if not exist "%SOLUTION_DIR%Fuwoa.AddIn\Fuwoa.snk" (
    sn -k "%SOLUTION_DIR%Fuwoa.AddIn\Fuwoa.snk" 2>nul
    if %ERRORLEVEL% NEQ 0 (
        echo ERROR: Failed to generate strong name key. Make sure you are running in Visual Studio Developer Command Prompt.
        pause
        exit /b 1
    )
    echo       Key generated: Fuwoa.snk
) else (
    echo       Key already exists, skipped
)

REM Step 2: Build solution
echo [2/4] Building solution...
msbuild "%SOLUTION_DIR%Fuwoa.sln" /t:Build /p:Configuration=%CONFIG% /p:Platform=x64

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo BUILD FAILED! Check the errors above.
    pause
    exit /b %ERRORLEVEL%
)
echo       Build succeeded.

REM Step 3: Locate the built DLL
echo [3/4] Locating built DLL...
set DLL_PATH=%SOLUTION_DIR%Fuwoa.AddIn\bin\%CONFIG%\Fuwoa.AddIn.dll

if not exist "%DLL_PATH%" (
    echo ERROR: DLL not found at %DLL_PATH%
    pause
    exit /b 1
)
echo       Found: %DLL_PATH%

REM Step 4: Register COM Add-In (requires admin)
echo [4/4] Registering COM Add-In (admin required)...
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe "%DLL_PATH%" /codebase /tlb

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo REGISTRATION FAILED! Please run this script as Administrator.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ============================================
echo  SUCCESS! Build and registration complete.
echo  Restart Excel to see the FUWOA tab.
echo ============================================
pause
