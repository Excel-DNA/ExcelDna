Excel-DNA - CSharpAddIn Sample
==============================

This sample creates a compiled library that will be loaded as an Excel add-in using the Excel-DNA runtime.

Items in the sample
-------------------
* CharpAddIn.sln - the solution file
* CharpAddIn.csproj - the project file
* Properties\AssemblyInfo.cs - project properties like version info
* Readme.txt - this readme file
* MyAddIn.cs - add-in source file with functions and macros
* SampleCS.dna - Excel-DNA directive file. 
  Copy to Output Directory: True

* Reference to ..\..\..\ExcelDna.Integration.dll
  Copy Local: False
* Post-build event:
  echo F | xcopy $(ProjectDir)..\..\..\ExcelDna.xll $(TargetDir)SampleCS.xll /C /Y

Target .NET Version
-------------------
* Target framework is .NET 4 in Project settings
* RuntimeVersion="v4.0" set in SampleCS.dna.

Debugging
---------
Breakpoints work best when I attach the debugger to the running process.
Do the following:
1. Build the project.
2. Start Excel externally.
3. Load the SampleCS.xll using File->Open.
4. In Visual Studio select Debug->Attach to Process,
   select Excel.exe
   select the debugger type (Attach to:...) to be "Managed (v4.0) code only"

To make changes in the code without restarting Excel:
1. Detach the debugger (Debug->Detach All).
2. Make changes and rebuild.
3. Reload the add-in in Excel (File->Open ... SamplesCS.xll).
4. Attach debugger again.