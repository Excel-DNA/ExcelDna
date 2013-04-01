This folder contains a variety of examples and experiments with the Excel-DNA async support.

1. AsyncFunctions folder - contains a C# project with the Reactive Extensions (RxExcel) examples, and some resizer experiments.

2. AsyncAwaitCS.dna - experiments with the C# 5 (.NET 4.5) async / await features.

3. AsyncMacros.dna - contains C# samples for ExcelAsyncUtil.RunAsMacro and some simple functions. Works 

4. ExcelTaskUtil.cs(vb) - Helper class in C#(VB) to integrate .NET 4 tasks with Excel-DNA async features.

5. ExcelTaskUtilTestCS(VB).dna - Sample add-in in C#(VB) using ExcelTaskUtil.

6. HttpClientVB.vb - Module used in ExcelTaskUtilTestVB.dna - implements async web fetching with cancellation using the System.Net.Http libraries (from the Microsoft.Net.Http package on NuGet).

7. FsAsync.dna - Samples in F# showing how the async workflows and native IObservable events can be hooked up to the Excel-DNA async support.
