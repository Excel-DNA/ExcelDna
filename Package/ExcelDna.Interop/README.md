# ExcelDna.Interop

This package provides local copies of these Microsoft Office primary interop assemblies:

- `Microsoft.Office.Interop.Excel.dll`
- `Microsoft.Vbe.Interop.dll`
- `office.dll`

The assemblies are available for .NET Framework 4.5.2 and `net6.0-windows7.0` projects. The package configures C# and Visual Basic projects to embed the interop types and not copy the assemblies to the output directory. F# projects are excluded from that build behavior.

For Excel-DNA documentation, see [excel-dna.net](https://excel-dna.net/).
