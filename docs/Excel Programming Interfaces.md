---
layout: page
title: Excel Programming Interfaces
---

Excel supports two completely different programming interfaces: 
1. The COM Automation interface that you know from VBA.
2. The native C API described in the Excel XLL SDK. 

Excel-DNA allows you to use either approach, and to mix them in your 
add-in to some extent. 

Excel-DNA integrates into Excel using the C API. The following types are related to the C API: 
* ExcelReference, which is a thin wrapper around a worksheet reference, and 
* XlCall.Excel and the XlCall constants, which give you a .NET interface based on the xlcall.h file that defines the C API. So whenever you're dealing with either an ExcelReference or calling Excel via the XlCall.Excel interface, you are dealing with the C API. Once place the ExcelReference type pops up, is if you want to receive a sheet reference as one of the arguments to a function. Normally you'd just get the value of the cell(s) as the argument passed in, so if you actually want the reference, you should set your parameter to by of type 'object' and add a special attribute, [ExcelArgument(AllowReference=true)](ExcelArgument(AllowReference=true)), to indicate that a reference should be passed. With this attribute, function calls that are passed a sheet reference will be called in your code with an object of type ExcelReference, allowing you to make further calls to the C API with this reference. 

On the other hand, the COM Automation interface can be used from your add-in as you would from VBA, taking the following into account: 
* To get hold of the root Application object for the Excel instance that is hosting your add-in, you should call the helper property "ExcelDnaUtil.Application". Once you have the root Application object, you can get hold of everything else from there. 
* To use the COM Automation interface you either need to use the C# 4 'dynamic' support, or reference an interop assembly that declares the COM types to .NET. 
* The support for calling COM was much improved in C# 4, whereas older versions you had to pass lots of "Missing" arguments, and properties weren't easy to work with. To run .NET 4 in your Excel-DNA add-in, you need to add a flag in the .dna file that sets RuntimeVersion="v4.0". 
* I normally recommend that one avoids using the COM Automation interface from within user-defined functions called from the worksheet. I have no good evidence, but suspect this sometimes causes 
issues. On the other hand, using the COM Automation interfaces elsewhere, like macros, ribbon handlers etc from your Excel-DNA seems to work perfectly. 