---
layout: page
title: Using the Excel COM Automation Interfaces
---

## Summary

To use the Excel COM object model from your macro or ribbon handler, you need to get hold of the right Application object. Use the {{ ExcelDnaUtil.Application }} call for this. If you assign this to a {{ dynamic }} variable (C# 4 only), or an 'Object' variable in VB.NET, then everything is straight-forward (but you have no IntelliSense help).

Otherwise you would reference an interop assembly that defines the COM types to your .NET project. You still get the right Application object from ExcelDnaUtil.Application, but now cast this to the type {{Microsoft.Office.Interop.Excel.Application}} and use it from there. 

There are different options for the interop assembly:

*  Use the version-specific official Primary Interop Assembly:
	* [Excel 2010 PIA](http://www.microsoft.com/en-us/download/details.aspx?id=3508)

* Install the "Excel-DNA.Interop" NuGet package, which gives you the Excel 2010 PIA assemblies.

* Use the version-independent [NetOffice](http://netoffice.codeplex.com) libraries.

If you use the "Embed Interop Types = True" support in .NET 4 (the default when referencing a PIA under .NET 4), you need not ship any assembly with your add-in.

## More info about Office versions and Interop Assemblies

Just like VBA code, as long as you stick to the parts of the COM object model that are common across versions, nothing in the COM interop part will be version-specific. This means if you you the Excel 2010 interop assemblies, but use only parts of the COM object model that are also there under Excel 2007, your add-in will work fine under Excel 2007 too. In that sense, the COM stuff is all version-independent. It's exactly the same as with VBA.

If you make a VBA add-in under Excel 2010, it will work under Excel 2003 too, as long as the parts of the object model you use are supported in both versions. In exactly the same way an Excel-DNA add-in that includes the NuGet "Excel-DNA.Interop" package (which gives you the Excel 2010 interop assemblies) will work fine under Excel 2003.  You can reference (and even distribute with your add-in) the interop assembly for Excel 2010, and it will work fine under Excel 2003, for those parts of the object model that are common.

The only problem you have regarding versions (both in VBA and making an Excel-DNA add-in), it that you can't see in the IntelliSense which parts of the COM object model are supported in which versions. So like in VBA, you might mistakenly use a method or property that does not exist in the hosting Excel version, and that will cause a runtime error. If this is a concern to you, then you can use the NetOffice (http://netoffice.codeplex.com) interop assemblies. They are version-independent, in the sense that they contain the union of all properties and methods from all Excel versions. In addition NetOffice gives you IntelliSense info that indicates under which Excel versions a particular method or property is supported. That's great, but the downside is that you have to distribute the NetOffice assemblies too, and can't just "Embed Interop Types" like with the Primary Interop Assemblies.

## COM reference management

In an Excel-DNA add-in, all access to the Excel COM object model should be from the main Excel thread. (A call to ExcelAsyncUtil.QueueAsMacro will allow you to initiate code running in a safe context on the main thread, from any other thread or context.)

When used only from the main thread, no special care is needed to manage the Excel COM object lifetimes. **Your Excel-DNA add-in should have no calls to `Marshal.ReleaseComObject` or `Marshal.FinalReleaseComObject`.** You also need not follow any rules related to the 'two-dot' access to COM objects. Code inside an Excel-DNA, running on the main thread, can safely access and use the Excel COM object model exactly like VBA code, with no additional concerns for reference counting.

Marshal.ReleaseComObject considered dangerous: [https://devblogs.microsoft.com/visualstudio/marshal-releasecomobject-considered-dangerous/](https://devblogs.microsoft.com/visualstudio/marshal-releasecomobject-considered-dangerous/)

Lifetimes of local variables under Debug vs. Release: [http://www.bryancook.net/2008/05/net-garbage-collection-behavior-for.html](http://www.bryancook.net/2008/05/net-garbage-collection-behavior-for.html)

## Samples

Using the .NET 4 'dynamic' type in C#:

{% highlight csharp %}
    [ExcelCommand(MenuName="Test", MenuText="Range Set")](ExcelCommand(MenuName=_Test_,-MenuText=_Range-Set_))
    public static void RangeSet()
    {
      dynamic xlApp = ExcelDnaUtil.Application;
      
      xlApp.Range["F1"](_F1_).Value = "Testing 1..2..3..4";
    }
{% endhighlight %}
