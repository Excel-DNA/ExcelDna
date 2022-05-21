setlocal

set PackageVersion=1.6.0-preview3
set DllVersion=1.6.0.0

set MSBuildPath="c:\Program Files\Microsoft Visual Studio\2022\Preview\Msbuild\Current\Bin\amd64\MSBuild.exe"

set rcfile=..\Source\versioninfo.rc2
PowerShell "(Get-Content %rcfile%) -replace '\d+,\d+,\d+,\d+', '%DllVersion%'.Replace('.',',') -replace '\d+\.\d+\.\d+\.\d+', '%DllVersion%' | Set-Content %rcfile%"
@if errorlevel 1 goto end

%MSBuildPath% ..\Source\ExcelDna.sln /t:restore /p:Configuration=Release
@if errorlevel 1 goto end

%MSBuildPath% ..\Source\ExcelDna.sln /t:Build /p:Configuration=Release
@if errorlevel 1 goto end

%MSBuildPath% ..\Source\ExcelDna.sln /t:Build /p:Configuration=Release /p:Platform=x64
@if errorlevel 1 goto end

call build.bat
@if errorlevel 1 goto end

cd ..\Package
call package.cmd %PackageVersion%
@if errorlevel 1 goto end

:end
