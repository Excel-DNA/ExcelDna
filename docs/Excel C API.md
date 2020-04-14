---
layout: page
title: Excel C API
---

The Excel C API was first introduced in Excel '95 and has continuously evolved with new versions of Excel. Excel-DNA uses the C API to integrate with Excel, though Excel-DNA add-in can also use the COM Automation interfaces where needed.

The Excel C API is documented in the Excel XLL Software Development Kit, available from Microsoft here: [http://msdn.microsoft.com/en-us/library/office/bb687883.aspx](http://msdn.microsoft.com/en-us/library/office/bb687883.aspx). The most coherent reference to using the C API is the book by Stephen Dalton - Financial Applications using Excel Add-in Development in C / C++.

Direct access to the C API is provided by Excel-DNA via two classes:
* `ExcelDna.Integration.ExcelReference`
* `ExcelDna.Integration.XlCall`

## ExcelReference
The `ExcelReference` class is a thin wrapper around the C API sheet reference information, and refers to a region (or multiple regions) on a specific sheet. ExcelReference has helper methods to get and set the values in the region, but for additional information, specific C API calls must be made.

Using C API calls it is possible to get the address of an ExcelReference - calling `XlCall.Excel(XlCall.xlfReftext, myReference, true)` - and from there a COM `Range` object may be contructed.

## XlCall
The header file in the Excel SDK that defines the C API is called XLCALL.H, so I've called the matching Excel-DNA class `XlCall`. It contains these functions:
* `XlCall.Excel` - Calls the `Excel4` / `Excel12` functions of the C API.
* `XlCall.TryExcel` - Same as `XlCall.Excel`, but returns an exact error result instead of throwing.
* `XlCall.RTD` - Wrapper around `XlCall.Excel(XlCall.xlfRtd, ...)` for the registration-free RTD support.

The first parameter to the `XlCall.Excel` / `XlCall.TryExcel` methods is a function code that identifies the function or command. These function codes are all available as constants on the `XlCall` class. The other parameters are marshaled by Excel-DNA from the .NET types to the corresponding C API types. (The parameter count passed in to the Excel function from C is not required in the Excel-DNA call.) Sheet references are passed to the C API as ExcelReference objects. Missing values (that are not at the end of the call) should be passed as ExcelMissing.Value.

Excel-DNA takes care of all type conversion and memory management for the C API calls.

## XlCall functions
There is no single source of documentation for all the XlCall functions. To piece together the available information, it helps to think of the functions in three classes:
* Auxiliary C API functions
* Worksheet functions
* Macro functions and commands

### Auxiliary C API functions
These are functions like `xlfRegister` and `xlSheetId` that can only be called from add-ins. These form part of the main Excel SDK and are documented in the Excel SDK documentation here: [http://msdn.microsoft.com/en-us/library/office/bb687870.aspx](http://msdn.microsoft.com/en-us/library/office/bb687870.aspx)

### Worksheet function
Most of the `XlCall.xlf...` function codes refer to regular Excel worksheet functions, like `xlfSum` and is documented in the regular Excel help.

### Macro functions and commands
Some information functions like `xlfGet...` and all the `xlc...` function codes refer to Excel 4.0 macro commands and macro sheet functions. These have no updated documentation, but are described in the Excel 4.0 macro help file. This file is officially available from Microsoft as a .hlp file here: [http://support.microsoft.com/kb/128185](http://support.microsoft.com/kb/128185) but is not easy to open in Windows 7/8. 
I also found a copy of this help file converted to a .chm file, kindly made available by Theo Heselmans here: [http://www.xceed.be/Blog.nsf/dx/excel-macro-function-help-file-for-windows-7](http://www.xceed.be/Blog.nsf/dx/excel-macro-function-help-file-for-windows-7). 
I've also made a  searchable .pdf version of the [Excel 4 Macro Reference](assets/Excel_C_API_Excel_4_Macro_Reference.pdf).


The help file documents the function in the macro form, but the names and parameters match the usage via the C API. For example, the macro function `GET.CELL(type_num, reference)` is called via the C API as `XlCall.Excel(XlCall.xlfGetCell, 7, myExcelReference)` - where the help file shows that 7 retrieves the number format of a cell. A comprehensive sample of the various `GET.XXX` information functions is available in the Excel-DNA distribution under Distribution\Samples\GetInfoAddIn.dna and the sheet GetInfoAddIn.xls.

Many of the C API functions have been discussed on the Excel-DNA Google group or elsewhere online, so searching at [http://groups.google.com/group/exceldna](http://groups.google.com/group/exceldna) is a good start.

### XlCall tricks
For add-ins made with VB.NET, the `XlCall` static class can be imported into a file with a call to 
`Imports ExcelDna.Integration.XlCall`
Then the functions can be called without the XlCall prefix:
`myResult = Excel(xlfGetCell, 7, myReference)`

Similarly for C#, a helper class that wraps some C API calls can derive from XlCall, as:
`static class MyHelpers : XlCall`

and functions inside the class can then call the Excel / TryExcel methods and all the XlCall constants without the XlCall prefix.