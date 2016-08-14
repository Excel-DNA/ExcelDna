@echo off
setlocal

set currentPath=%~dp0
set basePath=%currentPath:~0,-1%
set outputPath=%basePath%\nupkg

if exist "%outputPath%\*.nupkg" del "%outputPath%\*.nupkg"

if not exist "%outputPath%" mkdir "%outputPath%"

echo on
nuget.exe pack "%basePath%\ExcelDna.AddIn\ExcelDna.AddIn.nuspec" -BasePath "%basePath%\ExcelDna.AddIn" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe pack "%basePath%\ExcelDna.Integration\ExcelDna.Integration.nuspec" -BasePath "%basePath%\ExcelDna.Integration" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe pack "%basePath%\Excel-DNA\Excel-DNA.nuspec" -BasePath "%basePath%\Excel-DNA" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe pack "%basePath%\Excel-DNA.Lib\Excel-DNA.Lib.nuspec" -BasePath "%basePath%\Excel-DNA.Lib" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

:end