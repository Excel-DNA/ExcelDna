---
layout: post
title: "Async and event-streaming Excel UDFs with F#"
date: 2013-03-26 08:18:00 -0000
permalink: /2013/03/26/async-and-event-streaming-excel-udfs-with-f/
categories: features, samples, async, excel, exceldna, fsharp
---
There have been a some recent posts mentioning the [asynchronous][async-python] and [reactive][reactive] programming features in F#. Since Excel-DNA 0.30 added support for creating async and `IObservable`-based real-time data functions, I'd like to show how these F# features can be nicely exposed to Excel via Excel-DNA.

## IObservable to Excel via Excel-DNA

Excel-DNA 0.30 allows an add-in to expose `IObservable` sources to Excel as real-time data functions. (Actually Excel-DNA defines an interface called `IExcelObservable` that matches the semantics of `IObservable<object> - this is because we still target .NET 2.0 with the core library.)

Asynchronous function can then be implemented as an `IObservable` that returns a single value before completing. Cancellation (triggered when the user removes a real-time or async formula) is supported via the standard IDisposable mechanism.

Internally, Excel-DNA implements a thread-safe RTD server and registers the `IObservable` as an RTD topic. So some aspects of the `IObservable` support are subject to Excel's RTD feature works, for example the RTD throttle interval (by default 2 seconds) will also apply to `IObservable` functions.

The following functions in the `ExcelDna.Integration.ExcelAsyncUtil` helper class are relevant:

* `ExcelAsyncUtil.Initialize()` - this should be called in a macro context before any of the other features are used, typically from the `AutoOpen()` handler.

* `ExcelAsyncUtil.Observe(...)` - registers an IExcelObservable as a real-time data function with Excel. `Subsequent OnNext()` calls will raise updates via RTD.

* `ExcelAsyncUtil.Run(...)` - a helper method that runs a function asynchronously on a .NET threadpool thread.
 
In addition, we'll use

* `ExcelObservableSource` - a delegate type for functions that return an `IExcelObservable`.

### Some links:

* [Async functions in C#][async-cs-samples] - has some sample functions in C#.
* [Reactive Extensions for Excel (RxExcel)][rx-excel] - the RxExcel class is a small wrapper that bridges the `IExcelObservable` to any implementation of `IObservable<T>`, allowing the Rx libraries to be used in Excel.


## F# helpers for async and IObservable-based events

To initialize the async support in Excel-DNA, we need some code like the following:

{% highlight fsharp %}
namespace FsAsync

open System
open System.Threading
open System.Net
open Microsoft.FSharp.Control.WebExtensions
open ExcelDna.Integration

/// This class implements the IExcelAddin which allows us to initialize the ExcelAsyncUtil support.
/// It must not be a nested class (e.g. defined as a type inside a module) but a top-level class (inside a namespace)
type FsAsyncAddIn () =
    interface IExcelAddIn with
        member this.AutoOpen ()  = 
            ExcelAsyncUtil.Initialize ()
        member this.AutoClose () = ExcelAsyncUtil.Uninitialize ()

    // define a regular Excel UDF just to show that the add-in works
    [<ExcelFunction(Description="A test function from F#")>]
    static member fsaAddThem (x:double) (y:double) = x + y
{% endhighlight %}

F# supports an asynchronous programming model via "async computation expressions". The result of an async computation expression is a value of type `Async<T>`, which we need to convert to an `IExcelObservable`. We use a standard `CancellationTokenSource` hooked up to the `IDisposable` to enable cancellation.

{% highlight fsharp %}
module FsAsyncUtil =

    /// A helper to pass an F# Async computation to Excel-DNA 
    let excelRunAsync functionName parameters async =
        let obsSource =
            ExcelObservableSource(
                fun () -> 
                { new IExcelObservable with
                    member __.Subscribe observer =
                        // make something like CancellationDisposable
                        let cts = new CancellationTokenSource ()
                        let disp = { new IDisposable with member __.Dispose () = cts.Cancel () }
                        // Start the async computation on this thread
                        Async.StartWithContinuations 
                            (   async, 
                                ( fun result -> 
                                    observer.OnNext(result)
                                    observer.OnCompleted () ),
                                ( fun ex -> observer.OnError ex ),
                                ( fun ex ->
                                    observer.OnCompleted () ),
                                cts.Token 
                            )
                        // return the disposable
                        disp
                }) 
        ExcelAsyncUtil.Observe (functionName, parameters, obsSource)
{% endhighlight %}

Another neat feature of F# is that events are first-class types that implement `IObservable`. This means any F# event can serve as a real-time data source in an Excel formula. To bridge the F# events to the `IExcelObservable` interface is really easy, we just have the following helper:

{% highlight fsharp %}
/// A helper to pass an F# IObservable to Excel-DNA
let excelObserve functionName parameters observable = 
    let obsSource =
        ExcelObservableSource(
            fun () -> 
            { new IExcelObservable with
                member __.Subscribe observer =
                    // Subscribe to the F# observable
                    Observable.subscribe (fun value -> observer.OnNext (value)) observable
            })
    ExcelAsyncUtil.Observe (functionName, parameters, obsSource)
{% endhighlight %}


## Sample functions

Given the above helpers, we can now explore a few ways to implement async and real-time streaming functions. As examples:

Here is a plain synchronous function to download a url into a string:

{% highlight fsharp %}
let downloadString url = 
    try
        let uri = new System.Uri(url)
        let webClient = new WebClient()
        let html = webClient.DownloadString(uri)
        html
    with
        | ex -> "!!!ERROR: " + ex.Message
{% endhighlight %}

* **Async implementation 1**: Use Excel-DNA async directly to run `downloadString` on a `ThreadPool` thread

{% highlight fsharp %}
let downloadStringAsyncRunTP1 url = 
    ExcelAsyncUtil.Run ("downloadStringAsyncTP1", url, (fun () -> downloadString url :> obj))
{% endhighlight %}

Create an F# asynchronous computation for the download (this functions is not exported to Excel)

{% highlight fsharp %}
let downloadStringAsyncImpl url = async {
    try
        // In here we could check for cancellation using 
        // let! ct = Async.CancellationToken
        // if ct.IsCancellationRequested then ...
        let uri = new System.Uri(url)
        let webClient = new WebClient()
        let! html = webClient.AsyncDownloadString(uri)
        return html
    with
        | ex -> return "!!!ERROR: " + ex.Message 
    }
{% endhighlight %}

* **Async implementation 2**: This function runs the async computation synchronously on a `ThreadPool` thread because that's what `ExcelAsyncUtil.Run` does. Blocking calls will block a `ThreadPool` thread, eventually limiting the concurrency of the async calls

{% highlight fsharp %}
let downloadStringAsyncTP2 url = 
    ExcelAsyncUtil.Run ("downloadStringAsyncTP2", url, (fun () -> Async.RunSynchronously (downloadStringAsyncImpl url) :> obj))
{% endhighlight %}

* **Async implementation 3**: Use the helper we defined above. This runs the async computation using true F# async. Should not block `ThreadPool` threads, and allows cancellation

{% highlight fsharp %}
let downloadStringAsync url = 
    FsAsyncUtil.excelRunAsync "downloadStringAsync" url (downloadStringAsyncImpl url)
{% endhighlight %}

Helper that will create a timer that ticks at `timerInterval` for `timerDuration`, and is then done. Also not exported to Excel (incompatible signature). Notice that from F#, the `timer.Elapsed` event of the BCL `Timer` class implements `IObservable`, so can be used directly with the transformations in the F# `Observable` module.

{% highlight fsharp %}
let createTimer timerInterval timerDuration =
    // setup a timer
    let timer = new System.Timers.Timer(float timerInterval)
    timer.AutoReset <- true
    // return an async task for stopping it after the duration
    let timerStop = async {
        timer.Start()
        do! Async.Sleep timerDuration
        timer.Stop() 
        }
    Async.Start timerStop
    // Make sure that the type we actually observe in the event is supported by Excel
    // by converting the events to timestamps
    timer.Elapsed |> Observable.map (fun elapsed -> DateTime.Now) 
{% endhighlight %}

* **Event implementation**: Finally this is the Excel function that will tick away in a cell. Entered into a cell (_and formatted as a Time value_), the formula `=startTimer(5000, 60000)` will show a clock that ticks every 5 seconds for a minute.

{% highlight fsharp %}
let startTimer timerInterval timerDuration =
    FsAsyncUtil.excelObserve "startTimer" [|float timerInterval; float timerDuration|] (createTimer timerInterval timerDuration)
{% endhighlight %}


## Putting everything together in an Excel add-in

A complete `.dna` script file with the above code can be found in the [Excel-DNA distribution][exceldna-repo], under [Distribution\Samples\Async\FsAsync.dna][fsasync-sample].

Alternatively, the following steps would build an add-in in Visual Studio:

* Create a new F# library in Visual Studio.
* Install the Excel-DNA package from NuGet (`Install-Package Excel-DNA` from the NuGet console).
* Set up the Debug path:
  1. Select “Start External Program” and browse to find Excel.exe, e.g. for Excel 2010 the path might be: `C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE`.
  2. Enter the full path to the `.xll` file in the output as the Command line arguments, e.g. `C:\MyProjects\TestDnaFs\bin\Debug\TestDnaFs-addin.xll`.
* Place the following code in `Library1.fs`, compile and run:

{% highlight fsharp %}
namespace FsAsync

open System
open System.Threading
open System.Net
open Microsoft.FSharp.Control.WebExtensions
open ExcelDna.Integration

/// This class implements the IExcelAddin which allows us to initialize the ExcelAsyncUtil support.
/// It must not be a nested class (e.g. defined as a type inside a module) but a top-level class (inside a namespace)
type FsAsyncAddIn () =
    interface IExcelAddIn with
        member this.AutoOpen ()  = 
            ExcelAsyncUtil.Initialize ()
        member this.AutoClose () = ExcelAsyncUtil.Uninitialize ()

    // a regular Excel UDF just to show that the add-in works
    static member fsaAddThem (x:double) (y:double) = x + y

/// Some utility functions for connecting Excel-DNA async with F#
module FsAsyncUtil =
    /// A helper to pass an F# Async computation to Excel-DNA 
    let excelRunAsync functionName parameters async =
        let obsSource =
            ExcelObservableSource(
                fun () -> 
                { new IExcelObservable with
                    member __.Subscribe observer =
                        // make something like CancellationDisposable
                        let cts = new CancellationTokenSource ()
                        let disp = { new IDisposable with member __.Dispose () = cts.Cancel () }
                        // Start the async computation on this thread
                        Async.StartWithContinuations 
                            (   async, 
                                ( fun result -> 
                                    observer.OnNext(result)
                                    observer.OnCompleted () ),
                                ( fun ex -> observer.OnError ex ),
                                ( fun ex ->
                                    observer.OnCompleted () ),
                                cts.Token 
                            )
                        // return the disposable
                        disp
                }) 
        ExcelAsyncUtil.Observe (functionName, parameters, obsSource)

    /// A helper to pass an F# IObservable to Excel-DNA
    let excelObserve functionName parameters observable = 
        let obsSource =
            ExcelObservableSource(
                fun () -> 
                { new IExcelObservable with
                    member __.Subscribe observer =
                        // Subscribe to the F# observable
                        Observable.subscribe (fun value -> observer.OnNext (value)) observable
                })
        ExcelAsyncUtil.Observe (functionName, parameters, obsSource)

// Some test functions
module TestFunctions =
    /// Plain synchronous download function
    /// can be called from Excel
    let downloadString url = 
        try
            let uri = new System.Uri(url)
            let webClient = new WebClient()
            let html = webClient.DownloadString(uri)
            html
        with
            | ex -> "!!!ERROR: " + ex.Message

    /// Uses Excel-DNA async to run download on a ThreadPool thread
    let downloadStringAsyncTP1 url = 
        ExcelAsyncUtil.Run ("downloadStringAsyncTP1", url, (fun () -> downloadString url :> obj))

    /// Create an F# asynchronous computation for the download
    /// Not exported to Excel
    let downloadStringAsyncImpl url = async {
        try

            // In here we could check for cancellation using 
            // let! ct = Async.CancellationToken
            // if ct.IsCancellationRequested then ...
            let uri = new System.Uri(url)
            let webClient = new WebClient()
            let! html = webClient.AsyncDownloadString(uri)
            return html
        with
            | ex -> return "!!!ERROR: " + ex.Message 
        }

    /// This function runs the async computation synchronously on a ThreadPool thread
    /// because that's what ExcelAsyncUtil.Run does
    /// Blocking calls will block a ThreadPool thread, eventually limiting the concurrency of the async calls
    let downloadStringAsyncTP2 url = 
        ExcelAsyncUtil.Run ("downloadStringAsyncTP2", url, (fun () -> Async.RunSynchronously (downloadStringAsyncImpl url) :> obj))

    /// This runs the async computation using true F# async
    /// Should not block ThreadPool threads, and allows cancellation
    let downloadStringAsync url = 
        FsAsyncUtil.excelRunAsync "downloadStringAsync" url (downloadStringAsyncImpl url)

    // Helper that will create a timer that ticks at timerInterval for timerDuration, then stops
    // Not exported to Excel (incompatible type)
    let createTimer timerInterval timerDuration =
        // setup a timer
        let timer = new System.Timers.Timer(float timerInterval)
        timer.AutoReset  Observable.map (fun elapsed -> DateTime.Now) 

    // Excel function to start the timer - using the fact that F# events implement IObservable
    let startTimer timerInterval timerDuration =
        FsAsyncUtil.excelObserve "startTimer" [|float timerInterval; float timerDuration|] (createTimer timerInterval timerDuration)
{% endhighlight %}


## Support and feedback

The best place to ask any questions related to Excel-DNA is the [Excel-DNA Google group][excel-dna-group]. Any feedback from F# users trying out Excel-DNA or the features discussed here will be very welcome. I can also be contacted directly at <govert@icon.co.za>.

[async-python]: http://blogs.msdn.com/b/dsyme/archive/2013/03/24/asynchronous-programming-from-f-to-python.aspx
[reactive]: http://www.infoq.com/interviews/petricek-fsharp-functional-languages
[async-cs-samples]: http://exceldna.codeplex.com/wikipage?title=Asynchronous%20Functions
[rx-excel]: http://exceldna.codeplex.com/wikipage?title=Reactive%20Extensions%20for%20Excel
[exceldna-repo]: https://github.com/Excel-DNA/ExcelDna
[fsasync-sample]: https://github.com/Excel-DNA/ExcelDna/blob/master/Distribution/Samples/Async/FsAsync.dna
[excel-dna-group]: http://groups.google.com/group/exceldna
