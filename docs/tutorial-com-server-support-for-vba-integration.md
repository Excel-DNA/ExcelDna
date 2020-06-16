---
layout: post
title: "Tutorial: COM server support for VBA integration"
date: 2014-03-03 23:16:00 -0000
permalink: /2014/03/21/tutorial-com-server-support-for-vba-integration/
categories: uncategorized, .net, com, excel, excel-vba, exceldna, vba, xll
---
Functions and macros created in an Excel-DNA add-in can be called directly from Excel VBA by using `Application.Run(...)`. However, .NET also supports creating rich object models that are exported as COM libraries, which can be Tools->Referenced in VBA. Excel-DNA has some advanced support to host COM-exported objects from Excel-DNA add-ins, giving some advantages over the regular .NET "Register for COM interop" hosting approach:

* COM objects that are created via the Excel-DNA COM server support will be active in the same AppDomain as the rest of the add-in, allowing direct shared access to static variables, internal caches etc.

* COM registration for classes hosted by Excel-DNA does not require administrative access (even when registered via `RegSvr32.exe`).

* Everything needed for the COM server can be packed in a single-file .xll add-in, including the type library used for IntelliSense support in VBA.

[Mikael Katajamäki][mikael-katajamaki] has written some detailed tutorial posts on his [Excel in Finance][mikael-katajamaki] blog that explore this Excel-DNA feature, with detailed explanation, step-by-step instructions, screen shots and further links. See:

* [Interfacing C# and VBA with Excel-DNA (no intellisense support)][post-no-intellisense]
* [Interfacing C# and VBA with Excel-DNA (with intellisense support)][post-with-intellisense]

Note that these techniques would work equally well with code written in VB.NET, allowing you to port VB/VBA libraries to VB.NET with Excel-DNA and then use these from VBA.

Thank you Mikael for the great write-up!

[mikael-katajamaki]: http://mikejuniperhill.blogspot.com/
[post-no-intellisense]: http://mikejuniperhill.blogspot.com/2014/03/interfacing-c-and-vba-with-exceldna-no.html
[post-with-intellisense]: http://mikejuniperhill.blogspot.com/2014/03/interfacing-c-and-vba-with-exceldna_16.html
