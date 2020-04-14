---
layout: page
title: Excel - DNA Reference
---

## Helper classes/enums of ExcelDna.Integration

* ExcelReference - contains a sheet reference. Get/SetValue to read/write to the cells.
* ExcelError - an enum listing the different Excel errors
* ExcelDnaUtil - contains Application property (for COM access), WindowHandle and IsInFunctionWizard, * * ExcelVersion, ExcelLimits.

## Helper classes/enums of ExcelDna.Integration.CustomUI

* ExcelCommandBarUtil - has LoadCommandBars method to laod command bar xml from a string.
* CustomTaskPaneFactory - creates CustomTaskPanes.


**should we rather generate this from source code comments (XML, docfx)?**