setlocal

set PackageVersion=1.6.0-preview3
set DllVersion=1.6.0.0

set MSBuildPath="c:\Program Files\Microsoft Visual Studio\2022\Preview\Msbuild\Current\Bin\amd64\MSBuild.exe"

set rootPath=%~dp0..\..

set propsfile=%rootPath%\Directory.Build.props
copy /Y Directory.Build.props %propsfile%
PowerShell "(Get-Content %propsfile%) -replace '_VERSION_', '%DllVersion%' | Set-Content %propsfile%"
@if errorlevel 1 goto end

cd %rootPath%\ExcelDna\Build
call BuildPackages.bat %PackageVersion% %DllVersion% %MSBuildPath%
@if errorlevel 1 goto end

:end
