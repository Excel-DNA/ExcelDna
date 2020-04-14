---
layout: page
title: Asynchronous Functions with Tasks
---

This example shows how to implement an asynchronous Excel function in VB.NET using the .NET 4 Task class. This has an advantage over the ExcelAsyncUtil.Run method, which just runs the code on a ThreadPool thread. If we are able to use the .NET 4 Task class, the outstanding requests will not block any threads, so should scale better. Supporting .NET 4 Tasks also allow us to use the .NET 4.5 Async/Await language extensions.

We want to make an asynchronous function to download a string from a URL. I'm using the System.Net.Http package (if you're not using .NET 4.5, add to your project by getting the Microsoft.Net.Http package from NuGet). 

Our function, called _webDownloadString_ will be implemented like this:

{% highlight vbnet %}
Imports System.Net.Http
Imports ExcelDna.Integration

Public Module WebUdf
    ReadOnly myHttpClient As New HttpClient

    Public Function webDownloadString(url As String) As Object

        Return ExcelAsyncUtil.Observe("webDownloadString", url, Function() myHttpClient.GetStringAsync(url).ToExcelObservable())

    End Function
End Module
{% endhighlight %}

The implementation of the async function uses the ExcelAsyncUtil.Observe function, which takes an IExcelObservable as its last parameter. The HttpClient.GetStringAsync call returns a Task(Of String), so the missing part is the ToExcelObservable() function which converts a Task to an ExcelObservable.

ToExcelObservable is implemented like this:

{% highlight vbnet %}
Imports System.Threading.Tasks
Imports System.Runtime.CompilerServices
Imports ExcelDna.Integration

Module TaskExcelObservableExtensions

    ' Careful - this would only work as long as the task is not shared between calls, since cancellation cancels that task
    ' Another implementation would be via Reactive Extension: Task.ToObservable() (in System.Reactive.Linq.dll) and RxExcel
    <Extension()>
    Public Function ToExcelObservable(Of TResult)(task As Task(Of TResult)) As IExcelObservable

        If task Is Nothing Then
            Throw New ArgumentNullException("task")
        End If

        Return New ExcelTaskObservable(Of TResult)(task)
    End Function

    '' Wraps a Task in an Observable - basically allowing one Subscriber.
    Private Class ExcelTaskObservable(Of TResult)
        Implements IExcelObservable

        Private ReadOnly _task As Task(Of TResult)

        Public Sub New(task As Task(Of TResult))
            _task = task
        End Sub

        Public Function Subscribe(observer As IExcelObserver) As IDisposable Implements IExcelObservable.Subscribe
            Select Case _task.Status
                Case TaskStatus.RanToCompletion
                    observer.OnNext(_task.Result)
                    observer.OnCompleted()
                Case TaskStatus.Faulted
                    observer.OnError(_task.Exception.InnerException)
                Case TaskStatus.Canceled
                    observer.OnError(New TaskCanceledException(_task))
                Case Else
                    _task.ContinueWith(
                        Sub(t)
                            Select Case t.Status
                                Case TaskStatus.RanToCompletion
                                    observer.OnNext(t.Result)
                                    observer.OnCompleted()
                                Case TaskStatus.Faulted
                                    observer.OnError(t.Exception.InnerException)
                                Case TaskStatus.Canceled
                                    observer.OnError(New TaskCanceledException(t))
                            End Select
                        End Sub)
            End Select

            ' No cancellation
            Return DefaultDisposable.Instance
        End Function
    End Class

    Private Class DefaultDisposable
        Implements IDisposable
        Public Shared ReadOnly Instance As New DefaultDisposable()

        Private Sub New()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' no op
        End Sub
    End Class

End Module
{% endhighlight %}

We also need to initialize the Excel-DNA async support:
{% highlight vbnet %}
Imports ExcelDna.Integration

Public Class AddIn
    Implements IExcelAddIn

    Public Sub AutoOpen() Implements ExcelDna.Integration.IExcelAddIn.AutoOpen
        ExcelIntegration.RegisterUnhandledExceptionHandler(AddressOf HandleError)
        ExcelAsyncUtil.Initialize()
    End Sub

    Public Sub AutoClose() Implements ExcelDna.Integration.IExcelAddIn.AutoClose
        ExcelAsyncUtil.Uninitialize()
    End Sub

    Private Shared Function HandleError(ex As Object) As Object
        Return "!!! ERROR: " & ex.ToString()
    End Function
End Class
{% endhighlight %}

Note that the string returned would be truncated at the maximum string length for Excel - either 255 characters for Excel 2003, or 32767 characters for Excel 2007+.

The function can be called from an Excel sheet as **=webDownloadString("http://www.bing.com")**.
A next step might be to build a function that use regular expressions to extract data from the downloaded string.