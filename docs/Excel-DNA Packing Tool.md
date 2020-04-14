---
layout: page
title: Excel-DNA Packing Tool
---

Excel-DNA has a packing tool that allows you to create a single-file .xll add-in. 

When run on an add-in , the packing tool will compress and embed as resources into the .xll file the various parts of your add-in. The result is a single .xll file which contains your add-in and can be easily distributed. (The appropriate version of the .NET framework is still required on the client PC for your add-in to load.)

# Usage

The packing program is called _ExcelDnaPack.exe_ and is located in the _Distribution_ directory in the Excel-DNA download.

**`Usage: ExcelDnaPack.exe dnaPath [/O outputPath](_O-outputPath) [/Y](_Y)`**

:  dnaPath      The path to the primary .dna file for the ExcelDna add-in.
:  /Y           If the output .xll exists, overwrite without prompting.
:  /O outPath   Output path - default is <dnaPath>-packed.xll

## Preparation

Ensure that your add-in, including the required add-in assemblies and additional dependencies are working correctly from some directory. You might have the following files in _C:\FirstAddIn\_

- MyAddin.dna
- MyAddIn.xll _(a renamed copy of ExcelDna.xll from the distribution)_
- MyExcelFunctions.dll
- MyLib.dll

where _MyExcelFunctions.dll_ contains the add-in functions that are exported to Excel, and _MyLib.dll_ is a .NET assembly that is used by _MyExcelFunctions.dll_ but is not directly exported to Excel.

Update the .dna file to look like this

{% highlight xml %}
    <DnaLibrary Name="My First AddIn" RuntimeVersion="v4.0" >
        <ExternalLibrary Path="MyExcelFunctions.dll" Pack="true" />
        <Reference Path="MyLib.dll" Pack="true" />
    </DnaLibrary>
{% endhighlight %}

Note the following:
- I've added a _Pack="true"_ attribute to the ExternalLibrary tag, to indicate that this assembly should be packed.
- I've added an extra `<Reference Path=... Pack="true">` tag for the referenced assembly. The _Pack="true"_ directive tells ExcelDnaPack to embed the _MyLib.dll_ assembly into the .xll file and resolve the reference at runtime from this embedded resource.

One everything is in place, it's good to check that your add-in still works.

## Packing

Now run the packing program. From a command prompt, go to the C:\FirstAddIn\ directory and run _ExcelDnaPack.exe MyAddIn.dna_.
The output should look something like this:
![](assets/Excel-DNA_Packing_Tool_Packing_CommandLine_Output.PNG)

The result of the packing is a file called _MyAddIn-packed.xll_. This is the standalone .xll file, and can now be copied somewhere else and renamed to test.

## Other packing options

* If your add-in has a .xll.config file, say _MyAddIn.xll.config_, it will also be packed. Note that if an actual .xll.config file is present at runtime, it will be loaded instead of the embedded .xll.config file. This allows you to embed a default configuration, and still override it at runtime if required.

* Images can be embedded - useful for CommandBar and Ribbon custom UI extensions. Add tags
`<Image Name="ButtonImage" Path="MyButtonImage.png" Pack="true" />`
 and the image can be used from your ribbon .xml.

* .dna files can be nested, and will be embedded according to the Pack attributes:
`<ExternalLibrary Path="OtherFile.dna" Pack="true" />`
* SourceItems in Project tags can also be packed.

* The packing can also be incorporated into your build as a post-build step.

## Limitations

Mixed code assemblies (that contain both managed and native code) cannot currently be packed.

# Notes for 64-bit Excel

Note that a separate .xll must be created for the 64-bit version of Excel, using the _ExcelDna64.xll_ instead of _ExcelDna.xll_. For the packing, one option is to make and additional .dna file, say _MyAddIn64.dna_ and corresponding _MyAddIn64.xll_, which gets packed with _ExcelDnaPack.exe MyAddIn64.dna_. Then the final distribution will have two files:
- MyAddInFinal.xll (renamed from MyAddIn-packed.xll)
- MyAddInFinal64.xll (renamed from MyAddIn64-packed.xll)