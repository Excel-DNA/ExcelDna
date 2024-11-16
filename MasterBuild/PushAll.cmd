setlocal

set currentPath=%~dp0
set basePath=%currentPath:~0,-1%
set rootPath=%~dp0..\..


nuget.exe push "%rootPath%\ExcelDna\Package\nupkg\ExcelDna.Integration.1.9.0-alpha3.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\ExcelDna\Package\nupkg\ExcelDna.AddIn.1.9.0-alpha3.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\ExcelDna\Package\nupkg\Excel-DNA.Lib.1.9.0-alpha3.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\ExcelDna\Package\nupkg\Excel-DNA.1.9.0-alpha3.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\ExcelDna\Package\nupkg\ExcelDna.Templates.1.9.0-alpha3.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\IntelliSense\NuGet\nupkg\ExcelDna.IntelliSense.1.9.0-alpha3.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\ExcelDnaDoc\Package\nupkg\ExcelDnaDoc.1.9.0-alpha3.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\Registration\NuGet\nupkg\ExcelDna.Registration.1.9.0-alpha3.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\Registration\NuGet\nupkg\ExcelDna.Registration.FSharp.1.9.0-alpha3.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\Registration\NuGet\nupkg\ExcelDna.Registration.VisualBasic.1.9.0-alpha3.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\DeveloperTools\ExcelDna.Testing\Package\nupkg\ExcelDna.Testing.1.9.0-alpha3.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

:end
