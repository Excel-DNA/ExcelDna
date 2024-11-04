setlocal

set PackageVersion=%1
set DllVersion=%2
set MSBuildPath=%3

%MSBuildPath% ..\ExcelDnaN.sln /t:restore,build /p:Configuration=Release /p:Platform=x64 /p:ContinuousIntegrationBuild=true
@if errorlevel 1 goto end

rem call build.bat
@if errorlevel 1 goto end

rem cd ..\Package
rem call package.cmd %PackageVersion%
@if errorlevel 1 goto end

:end
