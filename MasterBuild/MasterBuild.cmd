setlocal

set PackageVersion=1.8.0
set PackageReferenceVersion=1.8.0
set DllVersion=1.8.0.5

set MSBuildPath="c:\Program Files\Microsoft Visual Studio\2022\Preview\Msbuild\Current\Bin\amd64\MSBuild.exe"

set rootPath=%~dp0..\..

set propsfile=%rootPath%\Directory.Build.props
copy /Y Directory.Build.props %propsfile%
PowerShell "(Get-Content %propsfile%) -replace '_VERSION_', '%DllVersion%' | Set-Content %propsfile%"
@if errorlevel 1 goto end

set targetsfile=%rootPath%\Directory.Build.targets.local
copy /Y PushAll.template.cmd PushAll.cmd
PowerShell "(Get-Content PushAll.cmd) -replace '_VERSION_', '%PackageVersion%' | Set-Content PushAll.cmd
@if errorlevel 1 goto end

set targetsfile=%rootPath%\Directory.Build.targets.local
copy /Y Directory.Build.targets %targetsfile%
PowerShell "(Get-Content %targetsfile%) -replace '_VERSION_', '%PackageReferenceVersion%' | Set-Content %targetsfile%"
@if errorlevel 1 goto end

set rootPathN=%~dp0..
cd %rootPathN%\Source\NativeAOT\Build
call BuildPackages.bat %PackageVersion% %DllVersion% %MSBuildPath%
@if errorlevel 1 goto end

:end
