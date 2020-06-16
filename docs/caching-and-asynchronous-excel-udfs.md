---
layout: post
title: "Caching and Asynchronous Excel UDFs"
date: 2013-04-02 23:32:00 -0000
permalink: /2013/04/02/caching-and-asynchronous-excel-udfs/
categories: samples, async, csharp, cache, excel, udf
---
This sample shows how the result of an Excel-DNA async UDF call can be cached using the .NET 4 MemoryCache class.

PS: Apparently there is a bug in the memory management of the .NET MemoryCache class. See the [StackOverflow discussion][memorycache-strangeness] and the [Connect bug report][memorycache-bug]. The [SharpMemoryCache NuGet package][sharp-memory-cache] might be an alternative, though I've not tried it.

{% highlight xml %}
<DnaLibrary Name="CachedAsyncSample" RuntimeVersion="v4.0" Language="C#">

  <Reference Name="System.Runtime.Caching" />
  <![CDATA[
    using System;
    using System.Threading;
    using System.Runtime.Caching;
    using ExcelDna.Integration;
    
    public static class dnaFunctions
    {
        public static object dnaCachedAsync(string input)
        {
            // First check the cache, and return immediately 
            // if we found something.
            // (We also need a unique key to identify the cache item)
            string key = "dnaCachedAsync:" + input;
            ObjectCache cache = MemoryCache.Default; 
            string cachedItem = cache[key] as string;
            if (cachedItem != null) 
                return cachedItem;
    
            // Not in the cache - make the async call 
            // to retrieve the item. (The second parameter here should identify 
            // the function call, so would usually be an array of the input parameters, 
            // but here we have the identifying key already.)
            object asyncResult = ExcelAsyncUtil.Run("dnaCachedAsync", key, () => 
            {
                // Here we fetch the data from far away....
                // This code will run on a ThreadPool thread.
    
                // To simulate a slow calculation or web service call,
                // Just sleep for a few seconds...
                Thread.Sleep(5000);
    
                // Then return the result
                return "The calculation with input " 
                        + input + " completed at " 
                        + DateTime.Now.ToString("HH:mm:ss");
            });
    
            // Check the asyncResult to see if we're still busy
            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
                return "!!! Fetching data";
    
            // OK, we actually got the result this time.
            // Add to the cache and return
            // (keeping the cached entry valid for 1 minute)
            // Note that the function won't recalc automatically after 
            //    the cache expires. For this we need to go the 
            //    RxExcel route with an IObservable.
            cache.Add(key, asyncResult, DateTime.Now.AddMinutes(1), null);
            return asyncResult;
        }
    
        public static string dnaTest()
        {
            return "Hello from CachedAsyncSample";
        }
    }
  
  ]]>
</DnaLibrary>
{% endhighlight %}

[memorycache-strangeness]: http://stackoverflow.com/questions/6895956/memorycache-strangeness
[memorycache-bug]: https://connect.microsoft.com/VisualStudio/feedback/details/806334/system-runtime-caching-memorycache-do-not-respect-memory-limits#
[sharp-memory-cache]: http://www.nuget.org/packages/SharpMemoryCache
