namespace ExcelDna.AddIn.RegistrationSampleFS
open System
open System.Threading
open ExcelDna.Integration

/// Some utility functions for connecting Excel-DNA async with F#
module FsAsyncUtil =

    // Ugly helper because typeof<FsAsyncUtil> won't work, since it's a module !???
    // http://stackoverflow.com/questions/2297236/how-to-get-type-of-the-module-in-f
    type internal Marker = interface end
    let ModuleType = typeof<Marker>.DeclaringType

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

    // Takes an Async and returns a one-shot IObservable
    let observeAsync async = 
        { new IObservable<_> with
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
        }
