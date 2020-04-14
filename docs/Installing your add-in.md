---
layout: page
title: Installing your add-in
---

Ease of deployment is one of the great advantages of making your Excel add-ins using Excel-DNA. This page provides some notes on the different approaches to installing your add-in.

The issue addressed here is how to install your add-in into Excel, so that it will automatically load every time Excel is started.

## Packing

First, you might find it useful to try the [Excel-DNA Packing Tool](Excel-DNA Packing Tool) to create a single .xll for deployment. 

## Security

Note that most Excel-DNA features do not require administrator rights to run, or any registration to be dome before running the add-in. This includes the use of RTD servers as well as Ribbon and Custom Task Pane UI customization.

Office implements a comprehensive security system, which is accessed via the _File->Options, Trust Center_ dialog. Your Excel-DNA add-ins are subject to the security restrictions configured in the Trust Center. To operate in a secure environment I suggest you sign the packed .xll (using the signtool utility) and incorporate the certificate into the Office Trust Center, along with the macro setting to allow only digitally signed macros to be loaded.

## Options to just run once

* Run from Visual Studio.  Set Project > Proprties > Debug to run Excel.exe (full path) and make .xll file the command line parameter.  (And use a post build task to copy and rename the ExcelDna.xll to the bin/debug folder.)
* Just manually File > Open the .xll file from Excel.
* VBA Application.RegisterXLL(...) can open it.  (Workbooks.Open will not work.)

## Options to install permanently

* In Excel, just File > Options > Addins.  Install as an ordinary addin, not as a COM add-in even though communication may be using COM.
* From VBA do
	* Application.addins.Add "myfilename.xll"
	* Application.addins("myTitle").Installed = True.  Title is defined by .dna file Name="myTitle", not necessarily the file name.
* Have an installer add registry
	* Key: HKCU/Software/Microsoft/Office/$version/Excel/Options, 
	* Values: OPEN, OPEN1, OPEN2 etc.) 
	* Set it to /R "C:\...\MyAddIn.xll"
	* (Do not leave gaps if uninstalling on Excel 2007+.  Eg OPEN, OPEN1, OPEN3.  Best to do nothing, Excel will sort itself out next time it is opened.)
	* (Can be done in a .bat file using reg)
* Install using COM, e.g. VB Script.  (But some sites ban vb script for being "Insecure".)

{% highlight vbnet %}
  Dim oXL As Object, oAddin As Object 
  Set oXL = CreateObject("Excel.Application") 
  oXL.Workbooks.Add 
  Set oAddin = oXL.AddIns.Add("C:\test.xll", False) 
  oAddin.Installed = True 
  oXL.Quit 
  Set oXL = Nothing 
{% endhighlight %}

* Another option is to install your add-in when it is opened the first time, by running some code in your AutoOpen macro. This way, the user only has to double-click your add-in the first time, and it will load and install, and load in future sessions. Your AutoOpen might be something like: 
{% highlight csharp %}
    public void AutoOpen() 
    { 
        string xllPath = (string)XlCall.Excel(XlCall.xlGetName); 
        var xlApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application; 
        xlApp.AddIns.Add(xllPath, false /**don't copy file**/).Installed = true; 
    } 
{% endhighlight %}

* Finally, it is possible to have ExcelDna be a separately installed COM server, and have Excel VBA access it using CreateObject etc.  There is an example in the samples, but this would be an uncommon approach.

## Creating a Windows Installer package for your add-in

* There's a fledgling [WiX-based installer project](https://github.com/Excel-DNA/WiXInstaller) available on GitHub - please help to improve it.
* Jiri Pik has written a [detailed guide](https://jiripik.com/2017/02/25/use-advanced-installer-excel-dna-project/) to creating an installer package using the commercial Advanced Installer tool.
