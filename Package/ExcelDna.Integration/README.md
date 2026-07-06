# ExcelDna.Integration

ExcelDna.Integration is the reference assembly for writing Excel-DNA add-ins. It contains the public attributes and helper APIs used by add-in code, including `ExcelFunction`, `ExcelArgument`, `ExcelCommand`, `ExcelDnaUtil`, and the ribbon/add-in interfaces.

If you are creating a complete `.xll` add-in, install `ExcelDna.AddIn` instead; it brings in this package and adds the build targets, loaders, and packing support. Install `ExcelDna.Integration` directly when you only need to compile against the Excel-DNA API, such as from a shared library.

```powershell
dotnet add package ExcelDna.Integration
```

Documentation and examples are available at https://excel-dna.net and https://github.com/Excel-DNA/ExcelDna.
