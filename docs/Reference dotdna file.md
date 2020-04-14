---
layout: page
title: Excel - DNA Reference: .dna file
---

## .dna file schema

Elements and attributes are below.   See *.dna files in downloaded samples.

* DnaLibrary 
	* Name
	* Description
	* Language
	* RuntimeVersion either "v4.0" for .Net 4, or "v2.0" (default) for _all_ earlier versions.  Ie. use "v2.0" for .Net 3.5.  Note that v4.0 is required for 64bit Office.
	* CompilerVersion
	* ShadowCopyFiles="true" enables shadow copying for the add-in's AppDomain.

* ExternalLibrary
	* Path
	* Pack
	* LoadFromBytes.  if ="true" then the .dll is loaded more dynamically which can avoid needing to shut down Excel during debugging.  However, it can cause issues with multiple active instances and harder debugging.
	* ExplicitExports="true" Prevents every static public function from becomming a UDF, they will need an explicit [ExcelFunction...](../Reference) Annotation.

* Project*
	* SourceItem*
* Image*

* CustomUI  Nest <customUI under this to define ribbon.


(Attributes set on the DnaLibrary apply to the implicit project and are not inherited by any <Project> sections.)

## Packing

ExcelDnaPack can pack the .dna files and dependent assemblies into a single .xll file.
The Samples\Packing directory has a number of samples of how to use the packing feature.
To run the samples, run PackAll.bat - the packed libraries are placed into the out directory.

## Default references and imports

A reference to the ExcelDna.Integration assembly in the .xll file that is 
loading a .dna file is always added when the .dna file is compiled.

In addition, the following references are added to each project, unless a DefaultReferences="false" attribute is set on the <Project>:
System.dll
System.Data.dll
System.Xml.dll
For RuntimeVersion="v4.0" the following are added too:
System.Core.dll
System.Data.DataExtensions.dll
System.Xml.Linq.dll
Microsoft.CSharp.dll (C# projects only)

For VB projects, the following imports are added, unless a DefaultImports="false" attribute is set on the <Project>:
Microsoft.VisualBasic
System
System.Collections
System.Collections.Generic
System.Data
System.Diagnostics
ExcelDna.Integration
For RuntimeVersion="v4.0" the following are added too:
System.Linq
System.Xml.Linq

?? Microsoft.Office.Core (for Ribbon)??

## Image Resolution

Image is only used if the Path is supplied.
Image Name must match (case sensitive) the image tag in the Ribbon markup.
1. If the Path starts with "packed:" the image is retrieved from resources.
2. If the Path contains .....

## Reference Resolution

The <Reference> element has Path (AssemblyPath is deprecated) and Name attributes. 
When compiling at runtime, the references are resolved as follows:
1. If Path starts with "packed:" the assembly is retrieved from resources, written to a temp file, and the temp file is added to the compiler's references list.
2. Else If Path is not null and not empty, the assembly is search relative to the .dna file and the xll directory, and in the framework directory. If found, Path is added to the compiler's references list.
3. Else the Name is passed to LoadWithPartialName for resolution.
4. Otherwise the reference is ignored.

When packing (Reference elements with the Pack="true" attribute), the reference is resolved as follows:
1. If Path starts with "packed:" the reference is ignored.
2. If Path is not null and not empty, the assemlby is looked for according to the Path Resolution rules.
3. If no file is found, Assembly.Load is attempted with the Path. If the load succeeeds, the assembly location is used to locate and pack the assembly.
4. Otherwise, Assembly.LoadWithPartialName with Name is tried. If the load succeeeds, the assembly location is used to locate and pack the assembly.
5. If the assembly is still not located, the reference is ignored for packing.

## Path Resolution

For assemblies and .dna files, the following path resolution is done. 
Given a path containing a file name, maybe rooted, maybe with some or no directory info, we attempt to find the file as follows:
1. If the file is found at the path, we are done.
2. Check the path - if it is rooted replace directory with .dna file's directory.
3. If the path is not rooted, try the whole path relative to the .dna file location (prepend the .dna file's directory).
4. Else try 2 or 3 using current AppDomain's BaseDirectory.


## ExplicitExports

If the ExplicitExports attribute is set to "true" on a Project or ExternalLibrary node, only functions and methods explicitly marked by an 
ExcelFunction or ExcelCommand attribute are exported. Otherwise, all public static methods with compatible signatures are exported.

## AutoOpen/AutoClose

Cleanup is only done when the add-in is removed from the add-in manager.
When File->Open is used to reopen the .xll, it is closed and opened, causing the .dna file to be re-read.