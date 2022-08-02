@echo off
setlocal

set currentPath=%~dp0
set basePath=%currentPath:~0,-1%
set outputPath=%basePath%\nupkg

if not exist "%outputPath%" mkdir "%outputPath%"

nuget.exe push "%outputPath%\Excel-DNA.Interop.15.0.0-preview3.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%outputPath%\ExcelDna.Interop.15.0.0-preview3.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%outputPath%\ExcelDna.Interop.Dao.15.0.0-preview3.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

:end
