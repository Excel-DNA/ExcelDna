# XML Schemas for Excel-DNA Add-Ins

This folder contains XML Schemas to help authoring of `.dna` files and other Office CustomUI elements such as Ribbon and Custom Task Panes (CTP), enabling IntelliSense and validation when editing the XML files, for example, in Visual Studio.

| File                                     | Description                                                                                                                                                                                          |
| ---------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| [`ExcelDna.DnaLibrary.xsd`][dna-xsd]     | XML Schema Reference for the `.dna` file format (`DnaLibrary`)                                                                                                                                       |
| [`customUI.xsd`][cui-2007-xsd]           | XML Schema Reference for Office 2007 Custom UI [provided by Microsoft][office-2007-xsd] and added here for convenience                                                                               |
| [`customui14.xsd`][cui-2010-xsd]         | XML Schema Reference for Office 2010 Custom UI [provided by Microsoft][office-2010-xsd] and added here for convenience                                                                               |
| [`ExcelDnaCatalog.xml`][dna-catalog-xml] | Mapping that tells Visual Studio to (automatically) use the `ExcelDna.DnaLibrary.xsd` XML Schema above for any files with the `.dna` extension. [See instructions below](#installation-instructions) |

[dna-xsd]: ExcelDna.DnaLibrary.xsd "XML Schema Reference for the `.dna` file format"
[cui-2007-xsd]: customUI.xsd "XML Schema Reference for Office 2007 provided by Microsoft"
[cui-2010-xsd]: customui14.xml "XML Schema Reference for Office 2010 provided by Microsoft"
[dna-catalog-xml]: ExcelDnaCatalog.xml "Mapping that tells Visual Studio to use the `ExcelDna.DnaLibrary.xsd` XML Schema for any files with the `.dna` extension"
[office-2007-xsd]: https://www.microsoft.com/en-us/download/details.aspx?id=4463 "2007 Office System: XML Schema Reference"
[office-2010-xsd]: https://www.microsoft.com/en-us/download/details.aspx?id=1574 "Office 2010 Reference: Office Fluent User Interface XML Schema"

## Installation instructions

There are different ways you can get IntelliSense and validation for Excel-DNA Add-Ins:

* [Install to a specific Excel-DNA project (via NuGet)](#install-to-a-specific-excel-dna-project-via-nuget) _(recommended)_
* [Install to a specific Excel-DNA project (manually)](#install-to-a-specific-excel-dna-project-manually)
* [Install globally to all projects on your machine (via VSIX)](#install-globally-to-all-projects-on-your-machine-via-vsix) _(soon)_
* [Install globally to all projects on your machine (manually)](#install-globally-to-all-projects-on-your-machine-manually)

### Install to a specific Excel-DNA project (via NuGet)

1. Install the NuGet package [`ExcelDna.XmlSchemas`](https://www.nuget.org/packages/ExcelDna.XmlSchemas/) on your Excel-DNA Add-In project:

    ```powershell
    install-package ExcelDna.XmlSchemas
    ```

    This package will add the 3 (three) XML Schema Definition files described above ([`ExcelDna.DnaLibrary.xsd`][dna-xsd], [`customUI.xsd`][cui-2007-xsd], and [`customui14.xsd`][cui-2010-xsd]) which Visual Studio will use to validate and provide IntelliSense to your `.dna` files and `.xml` files for Office Custom UI Ribbons, CTPs, etc.

2. Add the XML namespace `http://schemas.excel-dna.net/addin/2020/07/dnalibrary` to your `.dna` file(s). E.g.:

    ```xml
    <?xml version="1.0" encoding="utf-8"?>
    <DnaLibrary Name="Your Add-In" RuntimeVersion="v4.0" xmlns="http://schemas.excel-dna.net/addin/2020/07/dnalibrary">
      <!-- (...) -->
    </DnaLibrary>
    ```

### Install to a specific Excel-DNA project (manually)

1. Download the XML Schema Definition files below to a folder on your computer:

    * [`ExcelDna.DnaLibrary.xsd`][dna-xsd]
    * [`customUI.xsd`][cui-2007-xsd]
    * [`customui14.xsd`][cui-2010-xsd]

2. Add these files to your Excel-DNA Add-In project or solution, so that Visual Studio can detect them to perform validation and provide IntelliSense to your `.dna` files and `.xml` files for Office Custom UI Ribbons, CTPs, etc.

3. Add the XML namespace `http://schemas.excel-dna.net/addin/2020/07/dnalibrary` to your `.dna` file(s). E.g.:

    ```xml
    <?xml version="1.0" encoding="utf-8"?>
    <DnaLibrary Name="Your Add-In" RuntimeVersion="v4.0" xmlns="http://schemas.excel-dna.net/addin/2020/07/dnalibrary">
      <!-- (...) -->
    </DnaLibrary>
    ```

### Install globally to all projects on your machine (via VSIX)

(_coming soon_) follow this repository: <https://github.com/Excel-DNA/VSExcel>

### Install globally to all projects on your machine (manually)

This option requires you to copy files to Visual Studio's main installation folder, which means _you will need administrative privileges_ to perform get IntelliSense working with this approach.
In the instructions below, `%ProgramFiles(x86)%` is used to refer to the `Program Files` folder. This will typically be `C:\Program Files` on 32-bit systems and `C:\Program Files (x86)` on 64-bit systems. You may need to alter the paths below if you installed Visual Studio in a different location.

1. Download the XML Schema Definition files below to a folder on your computer:

    * [`ExcelDna.DnaLibrary.xsd`][dna-xsd]
    * [`customUI.xsd`][cui-2007-xsd]
    * [`customui14.xsd`][cui-2010-xsd]
    * [`ExcelDnaCatalog.xml`][dna-catalog-xml]

2. Copy these files to your the Visual Studio Global Schema Cache. This will be `%ProgramFiles(x86)%\Microsoft Visual Studio [xx]\Xml\Schemas\` for Visual Studio 2015 and earlier or `%ProgramFiles(x86)%\Microsoft Visual Studio\[xx]\[edition]\Xml\Schemas\` for Visual Studio 2017 and later. `[xx]` will be the version such as `14.0` or `2017` and `[edition]` will be the edition such as `Community`, `Professional`, or `Enterprise`.

    For more detailed information about [Visual Studio's Schema Cache][vs-schema-cache], see the documentation on Microsoft's website: [https://docs.microsoft.com/en-us/visualstudio/xml-tools/schema-cache][vs-schema-cache]

3. IntelliSense and validation should now work for all `.dna` files you open, automatically, even if they don't have the XML namespace for `.dna` files (`http://schemas.excel-dna.net/addin/2020/07/dnalibrary`) declared (although it a good practice to always include the XML namespace on your files). If it doesn't work, try restarting any open instances of Visual Studio and try again.

NOTE: Updates to Visual Studio may reset the global schema cache, undoing the changes made above. If that happens and you no longer see IntelliSense for `.dna` files, you'll need to repeat the steps above. Also, remember that other developers working on your project will not have IntelliSense or validation on their machines, unless they also perform the manual steps outlined above.

[vs-schema-cache]: https://docs.microsoft.com/en-us/visualstudio/xml-tools/schema-cache "Schema Cache"
