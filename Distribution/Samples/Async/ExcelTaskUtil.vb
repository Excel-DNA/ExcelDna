Imports System.Runtime.CompilerServices
Imports System.Threading
Imports System.Threading.Tasks
Imports ExcelDna.Integration

Public Module ExcelTaskUtil
    ' Another implementation via Reactive Extension is Task.ToObservable() (in System.Reactive.Linq.dll) with RxExcel
    <Extension()> _
    Public Function ToExcelObservable(Of TResult)(task As Task(Of TResult)) As IExcelObservable
        If task Is Nothing Then
            Throw New ArgumentNullException("task")
        End If

        Return New ExcelTaskObservable(Of TResult)(task)
    End Function

    <Extension()> _
    Public Function ToExcelObservable(Of TResult)(task As Task(Of TResult), cts As CancellationTokenSource) As IExcelObservable
        If task Is Nothing Then
            Throw New ArgumentNullException("task")
        End If

        Return New ExcelTaskObservable(Of TResult)(task, cts)
    End Function

    Public Function RunTask(Of TResult)(callerFunctionName As String, callerParameters As Object, taskSource As Func(Of CancellationToken, Task(Of TResult))) As Object
        Return ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, Function()
                                                                                Dim cts = New CancellationTokenSource()
                                                                                Dim task = taskSource(cts.Token)
                                                                                Return New ExcelTaskObservable(Of TResult)(task, cts)
                                                                            End Function)
    End Function

    Public Function RunTask(Of TResult)(callerFunctionName As String, callerParameters As Object, taskSource As Func(Of Task(Of TResult))) As Object
        Return ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, Function()
                                                                                Dim task = taskSource()
                                                                                Return New ExcelTaskObservable(Of TResult)(task)
                                                                            End Function)
    End Function

    Public Function RunAsTask(Of TResult)(callerFunctionName As String, callerParameters As Object, func As Func(Of CancellationToken, TResult)) As Object
        Return RunTask(callerFunctionName, callerParameters,
                       Function(cancellationToken) Task.Factory.StartNew(Of TResult)(Function() func(cancellationToken), cancellationToken))
    End Function

    Public Function RunAsTask(Of TResult)(callerFunctionName As String, callerParameters As Object, func As Func(Of TResult)) As Object
        Return RunTask(callerFunctionName, callerParameters, Function() Task.Factory.StartNew(Of TResult)(func))
    End Function

    ' Helper class to wrap a Task in an Observable - allowing one Subscriber.
    Private Class ExcelTaskObservable(Of TResult)
        Implements IExcelObservable
        ReadOnly _task As Task(Of TResult)
        ReadOnly _cts As CancellationTokenSource

        Public Sub New(task As Task(Of TResult))
            _task = task
        End Sub

        Public Sub New(task As Task(Of TResult), cts As CancellationTokenSource)
            Me.New(task)
            _cts = cts
        End Sub

        Public Function Subscribe(observer As IExcelObserver) As IDisposable Implements IExcelObservable.Subscribe
            Select Case _task.Status
                Case TaskStatus.RanToCompletion
                    observer.OnNext(_task.Result)
                    observer.OnCompleted()
                    Exit Select
                Case TaskStatus.Faulted
                    observer.OnError(_task.Exception.InnerException)
                    Exit Select
                Case TaskStatus.Canceled
                    observer.OnError(New TaskCanceledException(_task))
                    Exit Select
                Case Else
                    _task.ContinueWith(Sub(t)
                                           Select Case t.Status
                                               Case TaskStatus.RanToCompletion
                                                   observer.OnNext(t.Result)
                                                   observer.OnCompleted()
                                                   Exit Select
                                               Case TaskStatus.Faulted
                                                   observer.OnError(t.Exception.InnerException)
                                                   Exit Select
                                               Case TaskStatus.Canceled
                                                   observer.OnError(New TaskCanceledException(t))
                                                   Exit Select
                                           End Select
                                       End Sub)
                    Exit Select
            End Select

            ' Check for cancellation support
            If _cts IsNot Nothing Then
                Return New CancellationDisposable(_cts)
            End If
            ' No cancellation
            Return DefaultDisposable.Instance
        End Function
    End Class

    Private NotInheritable Class DefaultDisposable
        Implements IDisposable

        Public Shared ReadOnly Instance As New DefaultDisposable()

        ' Prevent external instantiation
        Private Sub New()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' no op
        End Sub
    End Class

    Private NotInheritable Class CancellationDisposable
        Implements IDisposable

        ReadOnly cts As CancellationTokenSource

        Public Sub New(cts As CancellationTokenSource)
            If cts Is Nothing Then
                Throw New ArgumentNullException("cts")
            End If

            Me.cts = cts
        End Sub

        Public Sub New()
            Me.New(New CancellationTokenSource())
        End Sub

        Public ReadOnly Property Token As CancellationToken
            Get
                Return cts.Token
            End Get
        End Property

        Public Sub Dispose() Implements IDisposable.Dispose
            cts.Cancel()
        End Sub
    End Class

End Module