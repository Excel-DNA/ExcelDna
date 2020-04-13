---
layout: page
title: Configure NLog logging
---
This is a guide to setting up logging from within your Excel-DNA add-in, using the NLog library. I used Visual Studio 2012 and C# to put together the example, but nothing here should be specific to the particular versions or language used.

## 1. Create a new C# Class Library
Start Visual Studio, select _New Project..._ and make a new C# Class Library project. I targeted the .NET Framework 4.5 since that is the default, but .NET Framework 4 should also work fine. I called my library **NLogTest**.

## 2. Install the Excel-DNA NuGet package
Open the NuGet Package Manager Console (Tools -> Library Package Manager -> Package Manager Console) and enter "Install-Package Excel-DNA". This will install the Excel-DNA package and display the Readme.txt file when complete.

## 3. Test basic add-in functionality
From inside the displayed Readme.txt file, copy the C# sample snippet into the file Class1.cs:

{% highlight csharp %}
using ExcelDna.Integration;

public static class MyFunctions
{
    [ExcelFunction(Description = "My first .NET function")](ExcelFunction(Description-=-_My-first-.NET-function_))
    public static string HelloDna(string name)
    {
        return "Hello " + name;
    }
}
{% endhighlight %}

(You can also rename the Class1.cs file if you want to.)

Notice that the NuGet package has set up the following aspects:
- Added a NLogTest-AddIn.dna file.
- Added a post-build step that will copy the ExcelDna.xll to the output directory as NLogTest-AddIn.xll.
- Added a post-build step that will run ExcelDnaPack to pack the add-in to produce NLogTest-AddIn-packed.xll in the output directory.
- Configured debugging to start Excel and load add-in.

Now press F5 to build and start debugging. This should load Excel and the add-in, and make the HelloDna function available. Try it in a cell as "=HelloDna("World!")

Exit Excel to stop debugging.

## Install the NLog library
Back in the Package Manager Console, install the NLog package: "Install-Package NLog". This should add a reference to the NLog library.

## Add a configuration file
The name of the actual Excel add-in is _NLogTest-AddIn.xll_ (this is the name set up by the Excel-DNA NuGet package). A configuration file for this add-in will be called _NLogTest-AddIn.xll.config_. We add such a file to the project: Right-click on the Project in the Solution Explorer, and select _Add -> New Item..._. Select _Application Configuration File_, and set the filename to _NLogTest-AddIn.xll.config_. This name must match the .xll name exactly (plus an extra .config extension). Right-click on the new file in the Solution Explorer, pick Properties, and set the file to: _Copy to Output Directory: Copy if newer_. This will ensure the file is copied to the output when changed.

## Add the NLog configuration entries
The configuration file might look like this:
{% highlight xml %}
<?xml version="1.0" encoding="utf-8" ?>
<configuration>
  <configSections>
    <section name="nlog" type="NLog.Config.ConfigSectionHandler, NLog"/>
  </configSections>
  <nlog xmlns="http://www.nlog-project.org/schemas/NLog.xsd" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <targets>
      <target name="logfile" xsi:type="File" fileName="${basedir}/LogFile.txt"/>
    </targets>
    <rules>
      <logger name="*" minLevel="Info" writeTo="logfile"/>
    </rules>
  </nlog>
</configuration>
{% endhighlight %}

## Add some logging to your function
{% highlight csharp %}
using ExcelDna.Integration;
using NLog;

public static class MyFunctions
{
    private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        
    [ExcelFunction(Description = "My first .NET function")](ExcelFunction(Description-=-_My-first-.NET-function_))
    public static string HelloDna(string name)
    {
        logger.Info("HelloDna function called: {0}", name);
        return "Hello " + name;
    }
}
{% endhighlight %}

Press F5 to build and start Excel, and enter the function into a cell: =HelloDna("World!")
(If there is an error in the configuration or in using NLog, the function will return _#VALUE!_)

Check the bin\Debug output directory  of the project - the LogFile.txt should be created, with the logged function call:
`2013-11-14 16:23:15.0465|INFO|MyFunctions|HelloDna function called: World!`

## Add NLog to the packed .xll
Excel-DNA can pack the dependencies of an .xll into a single file add-in. The _NLogTest-AddIn.xll.config_ file is automatically added by the ExcelDnaPack utility, but we also want to add the NLog library itself. To do this, update the _NLogTest-AddIn.dna_ file by adding a <Reference> entry for NLog as follows:

{% highlight xml %}
<DnaLibrary Name="NLogTest Add-In" RuntimeVersion="v4.0">
  <Reference Path="NLog.dll" Pack="true" />
  <ExternalLibrary Path="NLogTest.dll" LoadFromBytes="true" Pack="true" />
</DnaLibrary>
{% endhighlight %}


Then rebuild the project again.

## Check the final packed add-in
The resulting packed add-in, called _NLogTest-AddIn-packed.xll_ is a single-file add-in that includes the configuration file and NLog library. Copy this file from the bin\Debug output directory into a new directory (as the only file in the directory) and double-click to open in Excel. Then call the _HelloDna_ function again and check that the log file is created.

The resulting packed add-in can be renamed and distributed with no other dependencies required.
