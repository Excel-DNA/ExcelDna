---
layout: page
title: FSharp Standalone Assemblies
---

F# assemblies can be compiled with a -standalone switch that embeds the F# runtime into the assembly, and allows it to run without requiring additional assemblies.

There is a bug in Excel-DNA v 0.30 that prevents such assemblies to be loaded by Excel-DNA. The workaround is to:
- mark functions and macros that are to be exported with explicit `[<ExcelFunction>](_ExcelFunction_)` or `[<ExcelCommand>](_ExcelCommand_)` attributes, and
- set the 'ExplicitExports' flag on the ExternalLibrary tag in the .dna file:
`<ExternalLibrary Path="..." ExplicitExports="true" />`.

With these changes, F# libraries can be compiled with the standalone switch, and used as Excel-DNA add-ins without requiring additional F# libraries on the client.
