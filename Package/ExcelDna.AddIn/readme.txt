ExcelDna.AddIn - NuGet package for creating an Excel-DNA add-in
===============================================================
The Excel-DNA home page is at http://excel-dna.net.
For general support, please post to https://groups.google.com/forum/#!forum/exceldna.

Installing the ExcelDna.AddIn NuGet package into your project has made a number of changes to your project:

1. A file called <ProjectName>-AddIn.dna was added to your project, and set to be copied to the output directory. This is the master file for your Excel add-in.

2. A reference to <package>\lib\ExcelDna.Integration.dll was added. This contains helper classes like ExcelDnaUtil and ExcelFunctionAttribute that you may use in your add-in.

3. Under your Properties item group, a build properties file called ExcelDna.Build.props was added. This file allows build customization, like configuring whether the packing tool will be run.

4. Upon compilation, the project debugging settings will be configured to start Excel and load the appropriate (32-bit or 64-bit) unpacked version of your add-in, <ProjectName>-AddIn.xll or <ProjectName>-AddIn64.xll.

After building your project you will find (at least) the following files in your output directory (typically bin\Debug or bin\Release):

* <ProjectName>.dll & <ProjectName>.pdb - the normal outputs from compiling your library.
* <ProjectName>-AddIn.xll - the 32-bit native add-in loader for the unpacked add-in.
* <ProjectName>-AddIn.dna - a copy of the add-in master file for the 32-bit unpacked add-in.
* <ProjectName>-AddIn64.xll - the 64-bit native add-in loader for the unpacked add-in.
* <ProjectName>-AddIn64.dna - a copy of the add-in master file for the 64-bit unpacked add-in.
* <ProjectName>-AddIn-packed.xll - the 32-bit packed (all-in-one) redistributable add-in.
* <ProjectName>-AddIn64-packed.xll - the 64-bit packed (all-in-one) redistributable add-in.

For redistribution (if everything is set up correctly) you only need the two (32-bit and 64-bit) -packed.xll files. These files can be renamed as you like.

F# projects - special notes:
* You will need to build the project before attempting to debug (this ensures that the debug configuration is updated).
* F# projects built with newer Visual Studio versions should be configured to target .NET 4.5 or later, and ensure that the FSharp.Core.dll is copied to the output directory.
* Debugging will not be configured for F# projects when installing in Visual Studio 2013 or older. See instructions at the bottom of this file.

Next steps
----------
* Insert a sample function for your language from the Sample Snippets list below.
  Then press F5 to run Excel and load the add-in, and type into a cell: =HelloDna("your name")
* By default all Public Shared (public static in C#) functions (and functions in a Public Module) will be registered with Excel.
* Further configure packing for your library to add additional references by editing the <ProjectName>-Addin.dna file.
* To get IntelliSense and validation in your .dna files, you can install the NuGet package ExcelDna.XmlSchemas to add the .xml schema into the local project. Alternatively install the Visual Studio .vsix extension ExcelDna.XmlSchemas to add the schema file into Visual Studio. Further documentation here: https://github.com/Excel-DNA/ExcelDna/tree/master/Distribution/XmlSchemas/
* Source code, related projects and samples can be found on GitHub at https://github.com/Excel-DNA.
* Support questions at all levels are welcome at https://groups.google.com/forum/#!forum/exceldna.

Troubleshooting
---------------
Press F5 (Start Debugging) to compile the project, open the .xll add-in in Excel and make your functions available.

* If Excel does not open, check that the path under Project Properties->Debug->"Start external program:" is correct. If not, make sure that you have rebuilt the project successfully - this should automatically configure the debug options.
* If Excel starts but no add-in is loaded, check the Excel security settings under File -> Options -> Trust Center
  -> Trust Center Settings -> Macro Settings.
  Any option is fine _except_ "Disable all macros without notification."
* If Excel starts but you get a message saying "The file you are trying to open, [...], is in a different format than
  specified by the file extension.", then there is a mismatch between the bitness of Excel and the add-in being loaded.
* For any other problems, please post to the Excel-DNA group at https://groups.google.com/forum/#!forum/exceldna.

Uninstalling
------------
* When the ExcelDna.AddIn NuGet package is uninstalled, the <ProjectName>-AddIn.dna file will be renamed to
  "_UNINSTALLED_<ProjectName>-AddIn.dna" (to preserve any changes you've made). If the project won't be used as an Excel add-in this file may be deleted.

===============
Sample snippets
===============
Add one of the following snippets to your code to make your first Excel-DNA function.
Then press F5 to run Excel and load the add-in, and enter your function into a cell: =HelloDna("your name")
--------------
 Visual Basic
--------------

Imports ExcelDna.Integration

Public Module MyFunctions

    <ExcelFunction(Description:="My first .NET function")> _
    Public Function HelloDna(name As String) As String
        Return "Hello " & name
    End Function

End Module

----
 C#
----

using ExcelDna.Integration;

public static class MyFunctions
{
    [ExcelFunction(Description = "My first .NET function")]
    public static string HelloDna(string name)
    {
        return "Hello " + name;
    }
}

----
 F#
----

module MyFunctions

open ExcelDna.Integration

[<ExcelFunction(Description="My first .NET function")>]
let HelloDna name =
    "Hello " + name

---------------------------------------------------------------------
Configuring debugging in F# Projects with Visual Studio 2013 or older
---------------------------------------------------------------------
Debugging cannot be automatically configured for F# projects in older version of Visual Studio.
In the project properties, select the Debug tab, then
1. Select "Start External Program" and browse to find EXCEL.EXE, e.g. for Excel 2010 the path might
   be: C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE.
2. Enter the name to the .xll file in the output as the Command line arguments,
   e.g. "TestDnaFs-addin.xll"
        and for 64-bit Excel -addin64.xll.
