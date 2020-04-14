---
layout: page
title: Auto CLose and Excel Shutdown
---

This is a short note on the IExcelAddIn.AutoClose() callback, noting that it is not called when Excel is shut down, and explaining how the implementation came about.

Excel-DNA will call the IExcelAddIn.AutoClose() method when the add-in is removed from the add-ins dialog (Alt+t,i) by the user, or if the add-in is reloaded. In this case you can properly clean up your add-in - remove menus etc. Mostly when Excel shuts down you would not want to do a lot of clean-up - no need to remove menus, deregister functions etc. 

If you need to be notified of the Excel shutdown: 
* If you are running in Excel 2007+ and have an ExcelRibbon-derived 
class, just override the OnDisconnection or OnBeginShutdown. 
* To target any Excel version, add a new class that derives from ExcelComAddin, load it in your AutoOpen with ExcelComAddInHelper.LoadComAddIn(...), and override the OnDisconnection or OnBeginShutdown.

## Background

Excel .xll add-in export a few functions that are relevant to the discussion: 
* xlAutoOpen 
* xlAutoClose 
* xlAutoAdd 
* xlAutoRemove 

When an add-in is opened, xlAutoOpen is called, and Excel-DNA passes that trough to the IExcelAddIn.AutoOpen(). 

Excel calls xlAutoRemove when the add-in is removed from the Add-Ins dialog (thus if the user has explicitly chosen to remove the add-in from a running session). The problem is with xlAutoClose. If you have an add-in loaded in Excel with some workbook open and 'dirty', and then press Alt+F4, Excel will call the xlAutoClose, and then display a dialog to the user to ask whether to 'Save', 'Don't Save' or 'Cancel'. If the user selects 'Cancel' the session will continue. However, if the add-in has responded to the earlier xlAutoClose, it might now be removed although the session still continues, causing functions to fail and the add-in's ribbon to be missing. I didn't like this, so in Excel-DNA I only call  ExcelAddIn.AutoClose when I have received an xlAutoRemove before the xlAutoClose.

The resulting behaviour with Excel-DNA is that your add-in's AutoClose is only called when the add-in is actually removed by the user, and not when Excel exits. This allows AutoClose to do clean-up work that should be done when an add-in is removed. When Excel is shutting down, the whole process will be shut down, so your add-in should probably not do any clean-up. The operating system will close all handles, and recover all memory. Doing clean-up at this stage will just delay the closing of the Excel process. So I'm happy that this is a reasonable approach. 

In some cases the add-in might like to be notifies and do additional work when Excel is shutting down. Clearly xlAutoClose is not the right place for this, so the Excel C API does not give us an obvious way to implement such behaviour. We need some other mechanism to get a proper notification from Excel, and I suggest using the COM add-in approach, which works well. The COM add-in support in Excel-DNA which allows this is a much more recent addition, only implemented when I added support for the Excel 2007 ribbons, so was not an option when I initially decided what to do with AutoClose(). 

I have not added the COM add-in and it's shutdown event handler as a standard part of the add-in implementation, so that minimal Excel add-ins exposing only UDFs have no dependency on the COM support and so can claim to be 'pure .xll add-ins' using only the supported C API documented in the Excel SDK. In a sense, doing any COM stuff from the Excel-DNA add-in is making a hydrid with some hacks behind the scenes, and I think it is important to keep the COM part optional. 

Other events on the Excel Application object or the Workbook object might also be useful and hooked up from Excel-DNA, but there is no special support for these, apart from the ExcelDnaUtil.Application call which must be used to get hold of the correct Application root object. 

## Example Add-In

{% highlight csharp %}
<DnaLibrary RuntimeVersion="v4.0" Language="C#">
<Reference Name="System.Windows.Forms" />
<![CDATA[
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using SWF = System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.Extensibility;
using ExcelDna.Integration.CustomUI;

	[ComVisible(true)](ComVisible(true))
	public class MyComAddIn : ExcelComAddIn
    {
        public MyComAddIn()
        {
        }
        public override void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            SWF.MessageBox.Show("OnConnection");
        }
        public override void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            SWF.MessageBox.Show("OnDisconnection");
        }
        public override void OnAddInsUpdate(ref Array custom)
        {
            SWF.MessageBox.Show("OnAddInsUpdate");
        }
        public override void OnStartupComplete(ref Array custom)
        {
            SWF.MessageBox.Show("OnStartupComplete");
        }
        public override void OnBeginShutdown(ref Array custom)
        {
            SWF.MessageBox.Show("OnBeginShutDown");
        }
    }

    public class MyAddIn : IExcelAddIn
    {
        ExcelDna.Integration.CustomUI.ExcelComAddIn _comAddIn;

        public void AutoOpen()
        {
            try
            {
                _comAddIn = new MyComAddIn();
                ExcelComAddInHelper.LoadComAddIn(_comAddIn);
            }
            catch (Exception e)
            {
                SWF.MessageBox.Show("Error loading COM AddIn: " + e.ToString());
            }
        }

        public void AutoClose()
        {
        }
    }
]]>
</DnaLibrary>
{% endhighlight %}
