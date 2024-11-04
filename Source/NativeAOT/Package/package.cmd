@echo off
setlocal

set currentPath=%~dp0
set basePath=%currentPath:~0,-1%
set outputPath=%basePath%\nupkg
set ExcelDnaVersion=%1

if exist "%outputPath%\*.nupkg" del "%outputPath%\*.nupkg"

if not exist "%outputPath%" mkdir "%outputPath%"

echo on

nuget.exe pack "%basePath%\ExcelDna.AddInN\ExcelDna.AddInN.nuspec" -BasePath "%basePath%\ExcelDna.AddInN" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive -Prop ExcelDnaVersion="%ExcelDnaVersion%"
@if errorlevel 1 goto end

:end
