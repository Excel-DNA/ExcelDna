# ExcelDna.AddIn

Excel-DNA lets you build native Excel `.xll` add-ins with .NET. Add-ins can expose worksheet functions, macros, ribbon UI, RTD servers, custom task panes, and asynchronous functions while being distributed as ordinary Excel add-in files.

This is the main package for building a standard Excel-DNA add-in. It adds the Excel-DNA build targets, native 32-bit and 64-bit Excel loaders, the packing task, and a reference to `ExcelDna.Integration`.

```powershell
dotnet add package ExcelDna.AddIn
```

Add a function to your project:

```csharp
using ExcelDna.Integration;

public static class Functions
{
    [ExcelFunction(Description = "Returns a friendly greeting.")]
    public static string HelloDna(string name)
    {
        return "Hello " + name;
    }
}
```

Build the project, then load the generated `.xll` from Excel. The build output includes unpacked add-ins for debugging and packed add-ins for redistribution:

- `<ProjectName>-AddIn.xll` and `<ProjectName>-AddIn64.xll`
- `<ProjectName>-AddIn-packed.xll` and `<ProjectName>-AddIn64-packed.xll`

For most deployments, distribute the packed `.xll` files. They can be renamed to suit your product.

Excel-DNA supports .NET Framework 4.6.2 to 4.8.1 and modern .NET Windows targets. Documentation, examples, and troubleshooting notes are available at https://excel-dna.net and https://github.com/Excel-DNA/ExcelDna.
