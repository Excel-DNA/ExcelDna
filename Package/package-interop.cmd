@echo off
setlocal

set currentPath=%~dp0
set basePath=%currentPath:~0,-1%
set outputPath=%basePath%\nupkg
set ExcelDnaVersion=%1

if exist "%outputPath%\*.nupkg" del "%outputPath%\*.nupkg"

if not exist "%outputPath%" mkdir "%outputPath%"

echo on

nuget.exe pack "%basePath%\Excel-DNA.Interop\Excel-DNA.Interop.nuspec" -BasePath "%basePath%\Excel-DNA.Interop" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive -Prop ExcelDnaVersion="%ExcelDnaVersion%"
@if errorlevel 1 goto end

nuget.exe pack "%basePath%\ExcelDna.Interop\ExcelDna.Interop.nuspec" -BasePath "%basePath%\ExcelDna.Interop" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe pack "%basePath%\ExcelDna.Interop.Dao\ExcelDna.Interop.Dao.nuspec" -BasePath "%basePath%\ExcelDna.Interop.Dao" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

:end
