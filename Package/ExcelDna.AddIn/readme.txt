ExcelDna.AddIn - NuGet package for creating an Excel-DNA add-in
===============================================================
The Excel-DNA home page is at http://excel-dna.net.

Installing the ExcelDna.AddIn NuGet package into your project has made the following changes:
1. Added a reference to <package>\lib\ExcelDna.Integration.dll.
2. Added a build properties file called ExcelDna.Build.props under your Properties item group. This file allows build customization, and ensures that a copy of <package>\tools\ExcelDna\ExcelDna.xll will be placed in your output directory as 
   <ProjectName>-AddIn.xll. This is the add-in loader for your Excel add-in. 
3. The build targets file in the packages also lets the Excel-DNA packing tool ExcelDnaPack.exe run to create a single-file 
   redistributable, called <ProjectName>-AddIn-packed.xll.
4. Added a file called <ProjectName>-AddIn.dna to your project, set to be copied to the output directory (same name 
   as the .xll). This is the configuration file for your Excel add-in.
5. Configured <ProjectName>-Addin.dna to load your project library as an add-in library, and pack the compiled library
   into the redistributable.
6. Configured debugging to start Excel and load the <ProjectName>-AddIn.xll. 

   NOTE: Debugging will fail to load the add-in for 64-bit Excel installations. See below under Troubleshooting.
   NOTE: Debugging will not be configured for F# projects when installing in Visual Studio 2013 or older. See instructions at the bottom of this file.
   NOTE: F# projects built with newer Visual Studio versions should be configured to target .NET 4.5 or later, and ensure that the FSharp.Core.dll is copied to the output directory.

Next steps
----------
* Insert a sample function for your language from the Sample Snippets list below.
  Then press F5 to run Excel and load the add-in, and type into a cell: =HelloDna("your name")
* Add Public Shared functions (and functions in a Public Module) will be registered with Excel.
* Further configure packing for your library to add additional references by editing the <ProjectName>-Addin.dna file.
* Source code, related projects and samples can be found on GitHub at https://github.com/Excel-DNA.

Troubleshooting
---------------
Press F5 (Start Debugging) to compile the project, open the .xll add-in in Excel and make your functions available.

* If Excel does not open, check that the path under Project Properties->Debug->"Start external program:" is correct.
* If Excel starts but no add-in is loaded, check the Excel security settings under File -> Options -> Trust Center 
  -> Trust Center Settings -> Macro Settings. 
  Any option is fine _except_ "Disable all macros without notification."
* If Excel starts but you get a message saying "The file you are trying to open, [...], is in a different format than 
  specified by the file extension.", then you have the 64-bit version of Excel installed. Change the Debug command-line 
  parameters to start the "-AddIn64.xll" instead of "-AddIn.xll".
* For any other problems, please post to the Excel-DNA group at https://groups.google.com/group/exceldna.

Uninstalling
------------
* If the ExcelDna.AddIn NuGet package is uninstalled, the <ProjectName>-AddIn.dna file will be renamed to 
  "_UNINSTALLED_<ProjectName>-AddIn.dna" (to preserve any changes you've made) and may be deleted. 

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
