---
layout: page
title: Excel - DNA Reference: Wrapper SDK API, COM interface, Ribbon, Custom Task Panes, COM server support
---

## Wrapper for the Excel 97 / Excel 2007 SDK API

* XlCall class
XlCall.Excel wraps Excel4/Excel12 (but easy to call), also constants for all the API functions and commands.

* XlCallException - is thrown when the call to Excel fails.

* XlCall.TryExcel - does not throw exception on fail, but returns an XlCallReturn enum value.

* ExcelDna.Integration.Integration contains the static method 'RegisterMethods' which allow dynamic registration of methods.

## Excel COM interface access

ExcelDna.Integration.ExcelDnaUtils.Application returns the Excel Application COM object.
From VB this can be used late-bound quite easily.
From C# 4 late-binding through the dynamic 'type' is recommended.


## Ribbon
To support the Excel 2007 / 2010 Ribbon interface, the addin (in a .dna file or ExternalLibrary) must contain at least one public class that is a direct subclass of ExcelDna.Integration.CustomUI.ExcelRibbon.
This class can also implement ExcelDna.Integration.IExcelAddIn, but need not.
For each such class, ExcelDna will dynamically register and load a COM add-in in the AutoOpen call (after calling all the IExcelAddin.AutoOpen methods). This will trigger the loading of the Ribbon ui, and Excel calls the ExcelRibbon.GetCustomUI method. This is a virtual method, with a default implementation that retrieves the ribbon xml from the .dna file. An add-in can override the default GetCustomUI method to explicitly return an xml appropriate string. All callback methods that the Ribbon calls must be implemented as public methods in the class derived from ExcelRibbon.

My goal for the multi-version customUI support is to allow you to create a single add-in that contains UI customization for each version. The idea is not to make a unified customization layer - the add-in could contain different code for different versions.

## Custom Task Panes

Support under ExcelDna.Integration.CustomUI.
The CustomTaskPane class defines the interfaces related to CTP's.
A CTP must contain a UserControl (derived from System.Windows.Forms.UserControl).
Create a new CustomTaskPane containing an instance of MyUserControl by calling:
{% highlight csharp %}
    CustomTaskPane myCTP = 
        CustomTaskPaneFactory.CreateCustomTaskPane(typeof(MyUserControl), myTitle);
{% endhighlight %}


## COM server support

COM visible classes in ExternalLibrary tags marked ComServer='true', and COM visible classes that implement IRtdServer can be activated through the .xll directly. Even if the add-in is not loaded in Excel, such objects can be created in VBA.

These classes are (persistently) registered by calling "Regsvr32 <MyAddin>.xll" or by ComServer.DllRegisterServer(), and 
unregistered by "Regsvr32 /u <MyAddin>.xll or by ComServer.DllUnregisterServer().

Such classes can be accessed directly as RTD servers or from VBA using CreateObject("MyServer.ItsProgId"), and will be loaded in the add-in's AppDomain.
(The add-in need not be loaded for registered classes to be accessed through COM.)

A type library (.tlb) can be created for the assembly using tlbexp.exe, and will be registered if available. If the assembly is packed in the .xll, the type library will be packed too.

