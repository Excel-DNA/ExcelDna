---
layout: page
title: Excel - DNA Reference: Public types in ExcelDna library
---

Intended for use within user code (all in the namepace ExcelDna.Integration) are the following:

### Attributes 

* ExcelFunctionAttribute - for user-defined functions
	Name
	Description
	Category (by default the name of the add-in)
	HelpTopic
	IsVolatile (! suffix)
	IsMacroType (# suffix)
	IsThreadSafe (??? suffix)
	IsClusterSafe (& suffix)
* ExcelArgumentAttribute - for the arguments of user-defined functions
	Name
	Description
	AllowReference (R type) - Arguments of type object may receive ExcelReference.
* ExcelCommandAttribute - for macro commands
	Name
	Description
	HelpTopic
	ShortCut (does not seem to work ?)
	MenuName (default is library name)
	MenuText
	IsHidden
* ExcelComClassAttribute

**should we rather generate this from source code comments (XML, docfx)?**