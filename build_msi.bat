@echo off
setlocal
set MSBUILD=C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe
set WIX=C:\Program Files\WiX Toolset v7.0\bin\wix.exe
set ADDIN=%~dp0Fuwoa.AddIn\Fuwoa.AddIn.csproj
set CORE=%~dp0Fuwoa.Core\Fuwoa.Core.csproj
set WXS=%~dp0Fuwoa.Installer\Product.wxs
set DIST=%~dp0dist
set VERSION=1.0.2.0

echo === Building Fuwoa.Core ===
"%MSBUILD%" "%CORE%" /t:Build /p:Configuration=Release /p:Platform=x64 /nologo /v:m
if %errorlevel% neq 0 exit /b %errorlevel%

echo === Building Fuwoa.AddIn ===
"%MSBUILD%" "%ADDIN%" /t:Build /p:Configuration=Release /p:Platform=x64 /nologo /v:m
if %errorlevel% neq 0 exit /b %errorlevel%

echo === Building MSI: FUWOA_%VERSION%_x64.msi ===
if not exist "%DIST%" mkdir "%DIST%"
pushd "%~dp0Fuwoa.Installer"
"%WIX%" build --acceptEula wix7 -arch x64 -d "ProductVersion=%VERSION%" "Product.wxs" -o "%DIST%\FUWOA_%VERSION%_x64.msi"
popd
if %errorlevel% neq 0 exit /b %errorlevel%

echo === SUCCESS ===
dir "%DIST%\FUWOA_%VERSION%_x64.msi"
