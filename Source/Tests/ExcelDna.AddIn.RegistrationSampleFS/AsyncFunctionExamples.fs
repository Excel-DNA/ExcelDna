namespace ExcelDna.AddIn.RegistrationSampleFS

open System
open System.Threading
open System.Net
open Microsoft.FSharp.Control.WebExtensions
open ExcelDna.Integration

// Some test functions
module TestFunctions =
    /// Plain synchronous download function
    /// can be called from Excel
    let dnaFsDownloadString url = 
        try
            let uri = new System.Uri(url)
            #nowarn "FS0044"
            let webClient = new WebClient()
            #warnon "FS0044"
            let html = webClient.DownloadString(uri)
            html
        with
            | ex -> "!!!ERROR: " + ex.Message

    [<ExcelFunction>]
    let dnaFsHelloAsync name (msToWait: int) =
        async {
            do! Async.Sleep msToWait
            return "Hello " + name
        }
        |> FsAsyncUtil.observeAsync

    // Create an F# asynchronous computation for the download
    // Not normally exported to Excel (incompatible type)
    // but processed by AsyncRegistration helper
    [<ExcelFunction>]
    let dnaFsDownloadStringAsync url = async {
        try

            // In here we could check for cancellation using 
            // let! ct = Async.CancellationToken
            // if ct.IsCancellationRequested then ...
            let uri = new System.Uri(url)
            #nowarn "FS0044"
            let webClient = new WebClient()
            #warnon "FS0044"
            let! html = webClient.AsyncDownloadString(uri)
            return html
        with
            | ex -> return "!!!ERROR: " + ex.Message 
        }
                    
    // Create a timer that ticks at timerInterval for timerDuration, then stops
    // Not normally exported to Excel (incompatible type)
    // but processed by AsyncRegistration helper
    [<ExcelFunction>]
    let dnaFsCreateTimer timerInterval (timerDuration: int) =
        // setup a timer
        let timer = new System.Timers.Timer(float timerInterval)
        timer.AutoReset <- true
        // return an async task for stopping
        let timerStop = async {
            timer.Start()
            do! Async.Sleep timerDuration
            timer.Stop() 
            }
        Async.Start timerStop
        // Make sure that the type we observe in the event is supported by Excel
        // (events like timer.Elapsed are automatically IObservable in F#)
        timer.Elapsed |> Observable.map (fun elapsed -> DateTime.Now) 
