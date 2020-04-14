---
layout: page
title: Adding a keyboard shortcut
---

You can register a shortcut key for your macro in your AutoOpen. 

{% highlight csharp %}
<DnaLibrary Name="Test OnKey" Language="C#"> 
<![CDATA[ 
using ExcelDna.Integration; 
public class TestOnKey : IExcelAddIn 
{ 
    public void AutoOpen() 
    { 
        // Register Ctrl+Shift+H to call SayHello 
        XlCall.Excel(XlCall.xlcOnKey, "^H", "SayHello"); 
    } 
    public void AutoClose() 
    { 
        // Clear the registration if the add-in is unloaded 
        XlCall.Excel(XlCall.xlcOnKey, "^H"); 
    } 
        
    [ExcelCommand(MenuText = "Say Hello")](ExcelCommand(MenuText-=-_Say-Hello_)) 
    public static void SayHello() 
    { 
            XlCall.Excel(XlCall.xlcAlert, "Hello there!"); 
    } 
} 

]]> 
</DnaLibrary> 
{% endhighlight %}

This can also be done with the COM interface, using Application.OnKey.