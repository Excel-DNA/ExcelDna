using System;
using System.Diagnostics;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using ExcelDna.Integration;
using ExcelDna.Integration.RxExcel;

namespace AsyncTest
{
    public class AsyncRun
    {
        public static string HelloFromAsyncFunctions()
        {
            return "Hi there!";
        }

        public static object Sleep(string ms)
        {
            object result = ExcelAsyncUtil.Run("Sleep", ms, delegate
                {
                    Debug.Print("{1:HH:mm:ss.fff} Starting to sleep for {0} ms", ms, DateTime.Now);
                    Thread.Sleep(int.Parse(ms));
                    Debug.Print("{1:HH:mm:ss.fff} Completed sleeping for {0} ms", ms, DateTime.Now);
                    return "Woke Up at " + DateTime.Now.ToString("1:HH:mm:ss.fff");
                });
            if (Equals(result, ExcelError.ExcelErrorNA))
            {
                return "!!! Sleeping...";
            }
            return result;
        }


        public static object DownloadAsyncArray(string url)
        {
            Debug.Print("DownloadAsyncArray: " + url);
            // Don't do anything else here - might run at unexpected times...
            return ExcelAsyncUtil.Run("DownloadAsyncArray", url, delegate 
                { 
                    string result = Download(url);
                    string[,] resultArray = new string[,]
                    {
                        {"1: " + result},
                        {"2: " + result},
                        {"3: " + result}
                    };
                    return resultArray;
                });
        }


        public static object DownloadAsync(string url)
        {
            Debug.Print("DownloadAsync: " + url);
            // Don't do anything else here - might run at unexpected times...
            return ExcelAsyncUtil.Run("DownloadAsync", url, delegate { return Download(url); });
        }

        public static object WebSnippetAsync(string url, string regex)
        {
            Debug.Print("DownloadAsync: " + url);
            // Don't do anything else here - might run at unexpected times...
            return ExcelAsyncUtil.Run("WebSnippetAsync", new[] {url, regex}, delegate
                {
                    string result = Download(url);
                    string match = Regex.Match((result as string), regex, RegexOptions.Singleline).Groups[1].Value;

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

        public static object DownloadAsyncTask(string url)
        {
            return RxExcel.Observe("DownloadAsyncTask", url, () => new WebClient().DownloadStringTask(url));
        }

        // The 'parameters' can't change all the time. This function wil be called repeatedly forever, and return #N/A every time.
        public static object DownloadAsyncFail(string url)
        {
            Debug.Print("DownloadAsyncFail: " + url);

            return ExcelAsyncUtil.Run(
                "DownloadAsyncFail",
                DateTime.Now,       // <--- PROBLEM - parameters keep changing with every call.
                delegate { return Download(url); });
        }

    }
}

