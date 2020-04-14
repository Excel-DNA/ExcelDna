---
layout: page
title: User settings and the .xll.config file
---

1. Make a file called <TheAddInName>.xll.config with this in: 
{% highlight xml %}
<configuration> 
    <appSettings> 
        <add key = "Test" value="Forty-two" /> 
    </appSettings> 
</configuration> 
{% endhighlight %}

2. In your project, add a reference to the System.Configuration 
assembly. 

3. In your library add some function to access the settings: 
{% highlight csharp %}
internal static string GetAppSetting(string key) 
{ 
  object setting = 
System.Configuration.ConfigurationManager.AppSettings[key](key); 
  if (setting == null) 
  { 
    return "!! INVALID KEY !!"; 
  } 
  return setting.ToString(); 
} 
{% endhighlight %}

4. If you run ExcelDnaPack to pack the add-in into a single file, the .xll.config file will automatically be packed too. At runtime, if a .xll.config file is present, it will be used. Otherwise the packed .config file will be used as the configuration for for the add-in's AppDomain.