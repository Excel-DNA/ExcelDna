---
layout: page
title: Ribbon Customization
---

## Setting Ribbon Properties

The Ribbon extensibility model is a bit unusual. There is no opportunity to set the 'label' or 'image' property of the button after it is created, but there are 'getLabel' and 'getImage' callbacks that you can set up. 

To get Excel to refresh your control (or the whole Ribbon extension) you need to set an onLoad callback (on the customUI element) which receives an IRibbonUI interface for you to keep. This interface has two methods - Invalidate and InvalidateControl - which you call when a control should be refreshed. 

ExcelDna can helps with the implementation of the getImage callback - call the `ExcelRibbon.LoadImage` method (probably as `base.LoadImage(imageId)` in your code) with the imageId of the picture you want to show - this way you can load the images you specify in the .dna file. 

## References

I suggest the following references: 

### Articles on MSDN: 

* [Customizing the 2007 Office Fluent Ribbon for Developers (3 parts)](http://msdn.microsoft.com/en-us/library/aa338202(office.12).aspx)

* Making object wrappers to ease the scenario you have: [Custom Task Panes, the Office Fluent Ribbon, and Reusing VBA Code in the 2007 Office System](http://msdn.microsoft.com/en-us/library/bb194905(v=office.12).aspx)

#### Some Excel 2010 info: 

* [Customizing the backstage view in Excel 2010](http://msdn.microsoft.com/en-us/library/ee815851(v=office.14).aspx)

* [Customizing context menus in Excel 2010](http://msdn.microsoft.com/en-us/library/ee691832(v=office.14).aspx)

* [Tab activation and scaling in Excel 2010](http://msdn.microsoft.com/en-us/library/ee691834(v=office.14).aspx)

### Other links: 
* [Andy Pope's Ribbon Editor](http://www.andypope.info/vba/ribboneditor.htm) ([new, additional support for office 2010](http://www.andypope.info/vba/ribboneditor_2010.htm))

* [Ron de Bruin's site](http://www.rondebruin.nl/ribbon.htm) with details of the [Excel 2013 Backstage changes](http://www.rondebruin.nl/win/s2/win005.htm).

* [Discussion about Excel-DNA ribbons and the QAT](https://groups.google.com/forum/#!searchin/exceldna/qat/exceldna/hDocYCHy_Ao/SxnKUXDxiX8J).

Also note that a ribbon designed in VSTO can be exported to xml, which gives a <customUI ...> tag that can be used directly in Excel-DNA, though the ribbon handlers have to be re-implemented.