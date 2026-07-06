# ExcelDna.AddIn.NativeAOT

ExcelDna.AddIn.NativeAOT is preview support for publishing Excel-DNA add-ins with .NET Native AOT. It provides Native AOT-compatible Excel-DNA assemblies, a source generator, analyzers, native loader `.xll` files, and MSBuild targets that pack the published native output into an Excel add-in.

Use this package for Native AOT add-ins only. For ordinary managed Excel-DNA add-ins, use `ExcelDna.AddIn`.

```powershell
dotnet add package ExcelDna.AddIn.NativeAOT --prerelease
dotnet publish -c Release -r win-x64
```

The package supports Windows x64 and x86 runtime identifiers. The publish target selects the matching Excel-DNA loader and creates the add-in output after `Publish`.

Native AOT has stricter runtime and reflection rules than managed .NET. Keep the project AOT-friendly, test the published `.xll`, and avoid referencing the `ExcelDna.IntelliSense` package directly from a Native AOT add-in. To use IntelliSense with a Native AOT add-in, deploy the standalone ExcelDna.IntelliSense `.xll` alongside it.

See the Excel-DNA Native AOT documentation and samples in the Excel-DNA repository for the current preview guidance.
