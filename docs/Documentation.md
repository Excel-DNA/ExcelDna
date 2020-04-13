---
layout: page
title: Documentation
---

## Quick start

Get going with some first steps by following the [Getting Started](Getting Started) page.

To make a C# add-in with Visual Studio consult the [Step-by-step-CSharp-add-in.doc](assets/Step-by-step-CSharp-add-in.doc) guide.

Otherwise, browse through the list of other resources and examples on the [Documentation Home](index) page.

## How Tos
**Disclaimer:** Some of these are quite old and may not be conforming to the state of the current Excel-DNA version. There is an ongoing effort to update the existing and add new documentation.

* [Excel-DNA Packing Tool](Excel-DNA Packing Tool) The packing utility allow you to pack your add-in into a single .xll file for easy distribution.
* [Installing your add-in](Installing your add-in) and running generally.
* Accepting [Range Parameters](Range Parameters) in UDFs.
* [Integrating with VBA](Integrating with VBA) 
* [Performing Asynchronous Work](Performing Asynchronous Work)
* [Optional Parameters and Default Values](Optional Parameters and Default Values)
* [Keyboard Shortcut](Keyboard Shortcut)
* [Excel Programming Interfaces](Excel Programming Interfaces)
	* [Using the Excel COM Automation Interfaces](Using the Excel COM Automation Interfaces)
	* [Excel C API](Excel C API)
* [Ribbon Customization](Ribbon Customization) and various ribbon links.
* A note on [AutoClose and Detecting Excel Shutdown](AutoClose and Detecting Excel Shutdown).
* [Debugging Notes](Debugging notes)
* [COM Server Support](COM Server Support)
* Some notes on [FSharp Type Inference](FSharp Type Inference), and [FSharp Standalone Assemblies](FSharp Standalone Assemblies).
* [Asynchronous Functions](Asynchronous Functions)
* [Asynchronous Functions with Tasks](Asynchronous Functions with Tasks) example in VB.NET.
* [Reactive Extensions for Excel](Reactive Extensions for Excel)
* [Dynamic delegate registration](Dynamic delegate registration) - an advances feature to implement runtime registration and function wrappers.
* [User settings and the .xll.config file](User settings and the .xll.config file)
* A step-by-step guide to build a new add-in using the NuGet package, and then [Configure NLog logging](Configure NLog logging) for your add-in.
* [Creating a help file](Creating a help file)
* [Returning 1-D Arrays](Returning 1-D Arrays)
* [Async macro example - formatting the calling cell from a UDF](Async macro example - formatting the calling cell from a UDF)
* [Enumerating Excel COM Automation collections in VB.NET](Enumerating Excel COM Automation collections in VB.NET)
* [Modal dialog on new thread](Modal dialog on new thread)

### Reference
**Disclaimer:** The Reference section is still sparse and far from complete. Similar to above, there is an ongoing effort to complete the documentation.

* [Reference](Reference.md) - Main Page
* [Reference DataTypeMarshalling](Reference DataTypeMarshalling.md) - DataType treatment in Excel Functions
* [Reference dotdna file](Reference dotdna file.md) - Content of the .dna file
* [Reference HelperClasses](Reference HelperClasses.md) - Helper classes available in Excel-DNA
* [Reference PublicTypes](Reference PublicTypes.md) - available publich types in Excel-DNA
* [Reference Various](Reference Various.md) - Wrapper SDK API, COM interface, Ribbon, Custom Task Panes, COM server support

## Samples

[Distribution\Samples\](https://github.com/Excel-DNA/ExcelDna/tree/master/Distribution/Samples) contains various .dna files, each of which is a self-contained add-in that exhibit some Excel-DNA functionality.
The .dna files are .xml files that can be edited with a regular text editor.
To run any of the sample .dna files, make a copy of the Distribution\ExcelDna.xll file, place it next to the .dna file, and rename to have the same prefix. E.g. to run Optional.dna, make a copy of ExcelDna.xll called Optional.xll, and double-click, or File->Open to load in Excel.

## Support

There is a searchable record of more than 3460 messages on the Google group: [https://groups.google.com/group/exceldna](https://groups.google.com/group/exceldna.

**Please don't hesitate to ask.** If you are stuck or need some help using Excel-DNA your questions really are very welcome - whether you are just getting started, or an Excel-DNA expert.

And if you could help put together some proper documentation, please contact me. I'd be happy to add you as an editor on the Github project.

_-Govert_