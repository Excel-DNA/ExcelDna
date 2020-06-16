---
layout: post
title: "Excel UDF IntelliSense for Excel-DNA and VBA"
date: 2016-11-24 11:44:00 -0000
permalink: /2016/11/24/excel-udf-intellisense-for-excel-dna-and-vba/
categories: uncategorized
---
I'm happy to announce the first official release of the IntelliSense extension!

**Excel-DNA IntelliSense** provides on-sheet help for UDF functions as they are entered into a cell formula, similar to the help available for built-in Excel functions.

![Intellisense Release v1 01][intellisense-img1]
![Intellisense Release v1 02][intellisense-img2]

**For Excel-DNA add-ins** (v0.32 and later) that already provide descriptions in the `[ExcelFunction]` and `[ExcelArgument]` attributes, no extra work is needed. Just download and open (or install) the latest `ExcelDna.IntelliSense.xll` add-in from the GitHub ([https://github.com/Excel-DNA/IntelliSense/releases][intellisense-releases]), and the IntelliSense will light up. (There is also a NuGet package for embedding the service into your add-in, making distribution a bit easier.)
 
**For VBA functions**, you can add an extra sheet with the IntelliSense descriptions, or add an external .xml file with the information, or embed as a the `CustomXML` part in the Workbook or `.xlam` add-in.
Then open (or install) the `ExcelDna.IntelliSense.xll` add-in to provide the display service. Charles Williams, of [FastExcel][fastexcel] fame, has a detailed write-up on adding IntelliSense for your VBA function - see [https://fastexcel.wordpress.com/2016/10/07/writing-efficient-vba-udfs-part-15-adding-intellisense-to-your-udfs/][intellisense-vba].
 
**For PyXLL users**, the latest PyXLL 3.1 release offer built-in support for IntelliSense with the ExcelDna.IntelliSense.xll add-in installed. See [https://enthought.pyxll.com/whatsnew.html#intellisense][intellisense-pyxll].
 
Other native .xll add-ins can also provide IntelliSense through an external .xml file.
 
Details and downloads are on GitHub:
* Home: [https://github.com/Excel-DNA/IntelliSense][intellisense-repo]
* Releases: [https://github.com/Excel-DNA/IntelliSense/releases][intellisense-releases]
* Getting Started: [https://github.com/Excel-DNA/IntelliSense/wiki/Getting-Started][intellisense-getting-started]
* Detailed Usage Instructions: [https://github.com/Excel-DNA/IntelliSense/wiki/Usage-Instructions][intellisense-wiki] including details for incorporating the library into your own add-in for easier distribution.
 
**Public support and bug reports**:
The Excel-DNA Google group ([https://groups.google.com/forum/#!forum/exceldna][exceldna-google-group]) is the best place for general questions, comments etc. Detailed bug reports and feature requests can be added to the GitHub issues list: [https://github.com/Excel-DNA/IntelliSense/issues][intellisense-issues]
 
**Corporate support and private donations**:
If you find Excel-DNA and extensions like the IntelliSense service useful, please support the project by arranging a corporate support agreement, or making a donation via PayPal. See [https://excel-dna.net/support/][exceldna-support] for details and contact information.

[intellisense-img1]:  /assets/intellisense-release-v1-01.png "Intellisense Release v1 01"
[intellisense-img2]:  /assets/intellisense-release-v1-02.png "Intellisense Release v1 02"
[intellisense-releases]: https://github.com/Excel-DNA/IntelliSense/releases
[fastexcel]: http://www.decisionmodels.com/fastexcelD.htm
[intellisense-vba]: https://fastexcel.wordpress.com/2016/10/07/writing-efficient-vba-udfs-part-15-adding-intellisense-to-your-udfs/
[intellisense-pyxll]: https://enthought.pyxll.com/whatsnew.html#intellisense
[intellisense-repo]: https://github.com/Excel-DNA/IntelliSense
[intellisense-getting-started]: https://github.com/Excel-DNA/IntelliSense/wiki/Getting-Started
[intellisense-wiki]: https://github.com/Excel-DNA/IntelliSense/wiki/Usage-Instructions
[exceldna-google-group]: https://groups.google.com/forum/#!forum/exceldna
[intellisense-issues]: https://github.com/Excel-DNA/IntelliSense/issues
[exceldna-support]: /support/
