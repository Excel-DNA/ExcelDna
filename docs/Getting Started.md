---
layout: page
title: Getting Started with ExcelDna
---

## Do this first:
* The **.NET Framework 2.0 Runtime** or a later version must be installed. The .NET Framework Version 2.0 Redistributable Package is available from Microsoft.
* Get the most recent release of **ExcelDna**: Download [release:Excel-DNA version 1.00](https://github.com/Excel-DNA/ExcelDna/releases), unzip in a convenient directory.
* **Macro security** in Excel must not be 'Very High' or 'High' (setting to Medium is fine -- it will prompt whether to enable each macro library). To use the .NET macros you will have to 'Enable' at the prompt.

## 1. Create a user-defined function in Visual Basic
* Make a copy of ExcelDna.xll in a convenient directory, calling the copy Test1.xll.
* Create a new text file, called Test1.dna (the same prefix as the .xll file), with contents:

{% highlight vbnet %}
    <DnaLibrary>
    <![CDATA[

        Public Module MyFunctions

            Function AddThem(x, y)
                AddThem = x + y
            End Function

        End Module
    ]]>
    </DnaLibrary>
{% endhighlight %}

* Load Test1.xll in Excel (either _File->Open_ or _Tools->Add-Ins_ and _Browse..._).
* You should be prompted whether to Enable Macros, click Enable.
* Enter **=AddThem(4,2)** into a cell - you should get 6. (Under some localized versions of Excel the parameters are separated by a ';', so you'd say **=AddThem(4; 2)** instead).
* There should also be an entry for AddThem in the function wizard, under the category Test1.

Troubleshooting
* If you are not prompted to Enable Macros and nothing else happens, your security level is probably on High. Set it to Medium.
* If you get a message indicating the .Net 2.0 runtime could not be loaded, you might not have the .NET Framework 2.0 installed. Install it.
* If Excel crashes with an unhandled exception, an access violation or some other horrible error, either during loading or when running the function, please let me know. This shouldn't happen, and I would like to know if it does.
* If a window appears with the title 'ExcelDna Error Display' then there were some errors trying to compile the code in the .dna file. Check that you have put the right code into the .dna file.
* If Excel prompts for Enabling Macros, and then the function does not work and does not appear in the function wizard, you might not have the right filename for the .dna file. The prefix should be the same as the .xll file and it should be in the same directory.
* Excel is only able to load add-ins, including .xll add-ins, if VBA is installed. (This is under the Office Shared Tools option when installing Office.) Check that VBA is installed by pressing Alt+F11 to open the VBA editor. If it does not open, add VBA to your Office installation.
* Otherwise, if something goes wrong, let me know, or post on the discussion list.

## 2. Creating a user-defined function in C#
Change the contents of Test1.dna to:

{% highlight csharp %}
    <DnaLibrary Language="CS">
    <![CDATA[

        using ExcelDna.Integration;
	
                public class MyFunctions
                {
                        [ExcelFunction(Description="Joins a string to a number", Category="My functions")](ExcelFunction(Description=_Joins-a-string-to-a-number_,-Category=_My-functions_))
                        public static string JoinThem(string str, double val)
                        {
                                return str + val;
                        }
                }
    ]]>
    </DnaLibrary> 
{% endhighlight %}

* Reload the .xll, either from File->Open or in Tools->Add-Ins.
* Check with the formula =JoinThem("abc", 123)
* If the first example worked, this one should too.

## 3. Making the functions from a compiled library available
ExcelDna can also load any compiled .NET library. Public static functions with a compatible signature are exported to Excel.
* Create a file called TestLib.cs containing the following code:
{% highlight csharp %}
using ExcelDna.Integration;

public class MyFunctions
{
        [ExcelFunction(Description="Multiplies two numbers", Category="Useful functions")](ExcelFunction(Description=_Multiplies-two-numbers_,-Category=_Useful-functions_))
        public static double MultiplyThem(double x, double y)
        {
                return x * y;
        }
}
{% endhighlight %}
* You need to reference the ExcelDna.Integration.dll assembly (the ExcelFunction attribute is defined there). Copy the file ExcelDna.Integration.dll next to your source file, or reference it in your project. You need not redistribute this file - a copy is embedded as a resource in the redistributable .xll 
* Compile TestLib.cs to TestLib.dll: from the command-line: c:\windows\microsoft.net\framework\v2.0.50727\csc.exe /target:library /reference:ExcelDna.Integration.dll TestLib.cs
* Modify Test1.dna to contain:
{% highlight xml %}
        <DnaLibrary>
                <ExternalLibrary Path="TestLib.dll" />
        </DnaLibrary>
{% endhighlight %}
* Reload the .xll and check **=MultiplyThem(2,3)** (or **=MultiplyThem(2; 3)** on some versions).
* If you are compiling your assembly to target .NET 4, you need to tell Excel-DNA to load the right version of the CLR:
{% highlight xml %}
        <DnaLibrary RuntimeVersion="v4.0" >
                <ExternalLibrary Path="TestLib.dll" />
        </DnaLibrary>
{% endhighlight %}