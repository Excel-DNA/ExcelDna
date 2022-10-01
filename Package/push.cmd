@echo off
setlocal

set currentPath=%~dp0
set basePath=%currentPath:~0,-1%
set outputPath=%basePath%\nupkg

if not exist "%outputPath%" mkdir "%outputPath%"

nuget.exe push "%outputPath%\ExcelDna.Integration.1.6.0.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%outputPath%\ExcelDna.AddIn.1.6.0.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%outputPath%\Excel-DNA.Lib.1.6.0.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%outputPath%\Excel-DNA.1.6.0.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%outputPath%\ExcelDna.Templates.1.6.0.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end


:end
