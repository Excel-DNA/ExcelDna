ExcelDna.AddIn - NuGet package for creating an Excel-DNA add-in
===============================================================
The Excel-DNA home page is at http://excel-dna.net
For general support, please post to https://groups.google.com/g/exceldna
To encourage further development of the library, please visit https://github.com/sponsors/Excel-DNA


The ExcelDna.AddIn NuGet package was installed into your project in one of two ways:

1. Into the packages.config file in your project (the 'old' style package import).
  * A file called <ProjectName>-AddIn.dna was added to your project, and set to be copied to the output directory. This is the master file for your Excel add-in.
  * Under your Properties item group, a build properties file called ExcelDna.Build.props was added. This file allows build customization, like configuring whether the packing tool will be run.

2. As a <PackageReference> tag in your .csproj / .vbproj / .fsproj project file

  In this case the .dna file is not added to your project when the package in installed, but is created automatically in the output directory when the project builds.
  You can still have a .dna file in your project, but many settings from the .dna file can be added to the project file instead.


A reference to <package>\lib\ExcelDna.Integration.dll was also added. This contains helper classes like ExcelDnaUtil and ExcelFunctionAttribute that you may use in your add-in.

Upon compilation, the project debugging settings will be configured to start Excel and load the appropriate (32-bit or 64-bit) unpacked version of your add-in, <ProjectName>-AddIn.xll or <ProjectName>-AddIn64.xll.

After building your project you will find (at least) the following files in your output directory (typically bin\Debug or bin\Release):

* <ProjectName>.dll & <ProjectName>.pdb - the normal outputs from compiling your library.
* <ProjectName>-AddIn.xll - the 32-bit native add-in loader for the unpacked add-in.
* <ProjectName>-AddIn.dna - a copy of the add-in master file for the 32-bit unpacked add-in.
* <ProjectName>-AddIn64.xll - the 64-bit native add-in loader for the unpacked add-in.
* <ProjectName>-AddIn64.dna - a copy of the add-in master file for the 64-bit unpacked add-in.
* <ProjectName>-AddIn-packed.xll - the 32-bit packed (all-in-one) redistributable add-in.
* <ProjectName>-AddIn64-packed.xll - the 64-bit packed (all-in-one) redistributable add-in.

For redistribution (if everything is set up correctly) you only need the two (32-bit and 64-bit) -packed.xll files. These files can be renamed as you like.

Next steps
----------
* Insert a sample function for your language from the Sample Snippets list below.
  Then press F5 to run Excel and load the add-in, and type into a cell: =HelloDna("your name")
* By default all Public Shared (public static in C#) functions (and functions in a Public Module) will be registered with Excel.
* Source code, related projects and samples can be found on GitHub at https://github.com/Excel-DNA.
* Support questions at all levels are welcome at https://groups.google.com/forum/#!forum/exceldna.

Troubleshooting
---------------
Press F5 (Start Debugging) to compile the project, open the .xll add-in in Excel and make your functions available.

* If Excel does not open, check that the path under the 'Debug Properties' is correct. If not, make sure that you have rebuilt the project successfully - this should automatically configure the debug options.
* If Excel starts but no add-in is loaded, check the Excel security settings under File -> Options -> Trust Center
  -> Trust Center Settings -> Macro Settings.
  Any option is fine _except_ "Disable all macros without notification."
* If Excel starts but you get a message saying "The file you are trying to open, [...], is in a different format than
  specified by the file extension.", then there is a mismatch between the bitness of Excel and the add-in being loaded.
* For any other problems, please post to the Excel-DNA group at https://groups.google.com/g/exceldna

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
