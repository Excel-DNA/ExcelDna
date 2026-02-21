setlocal

set currentPath=%~dp0
set basePath=%currentPath:~0,-1%
set rootPath=%~dp0..\..


nuget.exe push "%rootPath%\ExcelDna\Package\nupkg\ExcelDna.Integration._VERSION_.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\ExcelDna\Package\nupkg\ExcelDna.AddIn._VERSION_.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\ExcelDna\Package\nupkg\ExcelDna.AddIn.NativeAOT._VERSION_.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\ExcelDna\Package\nupkg\Excel-DNA.Lib._VERSION_.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\ExcelDna\Package\nupkg\Excel-DNA._VERSION_.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\ExcelDna\Package\nupkg\ExcelDna.Templates._VERSION_.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\IntelliSense\NuGet\nupkg\ExcelDna.IntelliSense._VERSION_.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\ExcelDnaDoc\Package\nupkg\ExcelDnaDoc._VERSION_.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\Registration\NuGet\nupkg\ExcelDna.Registration._VERSION_.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\Registration\NuGet\nupkg\ExcelDna.Registration.FSharp._VERSION_.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\Registration\NuGet\nupkg\ExcelDna.Registration.VisualBasic._VERSION_.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe push "%rootPath%\DeveloperTools\ExcelDna.Testing\Package\nupkg\ExcelDna.Testing._VERSION_.nupkg" -Source  https://api.nuget.org/v3/index.json -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

:end
