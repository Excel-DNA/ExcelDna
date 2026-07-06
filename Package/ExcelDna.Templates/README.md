# ExcelDna.Templates

ExcelDna.Templates provides .NET project templates for starting Excel-DNA add-ins from the command line.

```powershell
dotnet new install ExcelDna.Templates
dotnet new excelfunc -lang C# -n MyFunctionsAddIn
dotnet new exceladdin -lang C# -n MyFullAddIn
```

The `excelfunc` template creates a small add-in with a sample worksheet function. The `exceladdin` template creates a fuller project with function, command, and ribbon examples. C#, F#, and Visual Basic templates are included.

After creating a project, build it and load the generated `.xll` in Excel. See https://excel-dna.net for more detail.
