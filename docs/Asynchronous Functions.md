---
layout: page
title: Asynchronous Functions
---

Excel-DNA now has a core implementation to support asynchronous functions. In a future version we might improve the ease of use.

## Usage

* You have to call {{ExcelAsyncUtil.Initialize()}} in your {{AutoOpen}}:

{% highlight csharp %}
    public class AsyncTestAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelAsyncUtil.Initialize();
            ExcelIntegration.RegisterUnhandledExceptionHandler(
                ex => "!!! EXCEPTION: " + ex.ToString());
        }

        public void AutoClose()
        {
        }
    }
{% endhighlight %}


* Your async UDF then calls {{AsyncUtil.Run}} like this:

{% highlight csharp %}
        public static object SleepAsync(string ms)
        {
            return ExcelAsyncUtil.Run("SleepAsync", ms, delegate
            {
                Debug.Print("{1:HH:mm:ss.fff} Sleeping for {0} ms", ms, DateTime.Now);
                Thread.Sleep(int.Parse(ms));
                Debug.Print("{1:HH:mm:ss.fff} Done sleeping {0} ms", ms, DateTime.Now);
                return "Woke Up at " + DateTime.Now.ToString("1:HH:mm:ss.fff");
            });
        }
{% endhighlight %}



The parameters to ExcelAsyncUtil.Run are:

	* {{string functionName}} - identifies this async function.
	* {{object parameters}} - identifies the set of parameters the function is being called with. Can be a single object (e.g. a string) or an object[]() array of parameters. It should include all the parameters to your UDF.
	* {{ ExcelFunc function}} - a delegate that will be evaluated asynchronously.

## More Samples

_ Note: This code does not scale very well, since the web calls block a treadpool thread. Using .NET 4 Tasks or .NET 4.5 async support could lead to a much better implementation. _

{% highlight csharp %}
        public static object DownloadAsync(string url)
        {
            // Don't do anything else here - might run at unexpected times...
            return ExcelAsyncUtil.Run("DownloadAsync", url,
                delegate { return Download(url); });
        }

        public static object WebSnippetAsync(string url, string regex)
        {
            // Don't do anything else here - might run at unexpected times...
            return ExcelAsyncUtil.Run("WebSnippetAsync", new object[]() {url, regex},
                delegate
                {
                    string result = Download(url);
                    string match = Regex.Match((result as string), regex,          
                        RegexOptions.Singleline).Groups[1](1).Value;

                    match = Regex.Replace(match, "\r", " ");
                    match = Regex.Replace(match, "\n", " ");
                    match = Regex.Replace(match, "\t", " ");
                    return match;
                });
        }

        static string Download(string url)
        {
            return new WebClient().DownloadString(url);
        }
{% endhighlight %}

