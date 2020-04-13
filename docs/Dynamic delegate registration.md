---
layout: page
title: Dynamic delegate registration
---

In come cases one might want to implement some kind of function wrapper or transformation at runtime. E.g. automatically wrapping and registering async Task / Rx functions.

The latest check-ins for Excel-DNA (check-in 79681 and eventually version 0.32) implement the support required to do this.

The key method that has been added is {{ExcelIntegration.RegisterDelegates(…)}}, which allows you to pass in a list of delegates, together with lists of {{ExcelFunction}} / {{ExcelArgument}} attributes. Because this takes {{Delegates}} and not just {{MethodInfos}}, you can easily wrap an existing method with to include your processing code for the optional / default values.

A useful helper that complements this is {{ExcelIntegration.GetExportedAssemblies()}} which returns the {{Assemblies}} that were considered for registration by Excel-DNA - either from {{ExternalLibrary}} tags in the .dna file or from runtime-compiled source projects inside the .dna file.

The basic idea would be:

In your {{AutoOpen}}, call some kind of {{UpdateRegistrations()}} which works like this:
1.	Get all the methods you’re interested in via Reflection (from the assemblies returned by {{ExcelIntegration.GetExportedAssemblies()}}).
2.	Build delegates using lambda expressions that add the optional handling (or using the Expression Tree API for even more control).
3.	Register the delegates with the right attributes via {{ExcelIntegration.RegisterDelegates}}.


This code should be a start:

{% highlight csharp %}
<DnaLibrary Name="Dynamic Function Tests" Language="C#" RuntimeVersion="v4.0">
<Reference Name="System.Windows.Forms" />
<![CDATA[
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Windows.Forms;
using ExcelDna.Integration;

public class TestAddIn : IExcelAddIn
{
    public void AutoOpen() 
    { 
        try
        {
            MessageBox.Show("In AutoOpen");
        
            var helloDel = MakeDelegate("Hello ");
            var byeDel = MakeDelegate("Goodbye ");

            var helloAtt = new ExcelFunctionAttribute
            {
                Name = "delHello",
            };
            var helloArgAtt = new ExcelArgumentAttribute
            {
                Name = "theName",
                Description = "is the name of the person to say 'Hello' to."
            };
              
            var byeAtt = new ExcelFunctionAttribute
            {
                Name = "delGoodbye",
            };
            var byeArgAtt = new ExcelArgumentAttribute
            {
                Name = "theName",
                Description = "is the name of the person to say 'Goodbye' to."
            };
              
            var add3Del = MakeAddNumber(3);
            var add3Att = new ExcelFunctionAttribute
            {
                Name = "delAdd3",
                Description = "Adds 3 to a number",
                IsThreadSafe = true,
                IsExceptionSafe = true
            };
            var add3ArgAtt = new ExcelArgumentAttribute
            {
                Name = "theNumber",
                Description = "is the number to which the adding is done."
            };
              
            ExcelIntegration.RegisterDelegates(
              new List<Delegate> { helloDel, byeDel, add3Del }, 
              new List<object>   { helloAtt, byeAtt, add3Att },
              new List<List<object>> { new List<object> {helloArgAtt}, 
                                       new List<object> {byeArgAtt},
                                       new List<object> {add3ArgAtt},
                                     } );
        }
        catch (Exception ex)
        {
              MessageBox.Show(ex.ToString());
        }
    } 

	public void AutoClose() {}
    
    static Func<string, string> MakeDelegate(string sayWhat)
	{
		Func<string, string> saySomethingToName = name => sayWhat + name;
		return saySomethingToName;
	}
	
	static Func<double, object> MakeAddNumber(double numberToAdd)
	{
	  return x => 
	  {
		try
		{
		  return x + numberToAdd;
		}
		catch (Exception ex)
		{
		  return double.NaN;
		}
	  };
	}
}

]]>
</DnaLibrary>
{% endhighlight %}
