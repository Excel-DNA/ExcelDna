---
layout: page
title: Reactive Extensions for Excel - VB.NET
---

{% highlight vbnet %}
Imports System.Runtime.CompilerServices
Imports ExcelDna.Integration

Public Module RxExcel

    <Extension()>
    Public Function ToExcelObservable(Of T)(observable As IObservable(Of T)) As IExcelObservable
        Return New ExcelObservable(Of T)(observable)
    End Function

    Public Function Observe(Of T)(functionName As String, parameters As Object, _
                           observableSource As Func(Of IObservable(Of T))) As Object
        Return ExcelAsyncUtil.Observe(functionName, parameters, 
                                     Function() observableSource().ToExcelObservable())
    End Function
End Module

Public Class ExcelObservable(Of T)
    Implements IExcelObservable

    ReadOnly _observable As IObservable(Of T)

    Public Sub New(observable As IObservable(Of T))
        _observable = observable
    End Sub

    Public Function Subscribe(observer As IExcelObserver) As IDisposable _
        Implements IExcelObservable.Subscribe
        Return _observable.Subscribe(Sub(value) observer.OnNext(value), 
            Sub(ex) observer.OnError(ex), Sub() observer.OnCompleted())
    End Function
End Class
{% endhighlight %}