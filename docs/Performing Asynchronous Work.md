---
layout: page
title: Performing Asynchronous Work
---

It is important to only communicate with Excel when it is Ready.   For example, if one displays a non modal dialog then the program may fail if it attempts to call a VBA routine in response to a button press while the user is also in the process of editing a cell.  Intercepting Window Messages messages etc. can also produce errors even though they are on the main thread.   And of course accessing Excel from a different thread is forbidden.

The best approach is to enqueue such work to execute on the main Excel thread when it is ready.  Some support for initiating such cross-thread work is now (December 2012) implemented by Excel-DNA, and exposed as methods on the `ExcelDna.Integration.ExcelAsyncUtil` class.

To try it you need to 
- call `ExcelAsyncUtil.Initialize()` in your `AutoOpen()`. 
- when you want Excel to do the work, call `ExcelAsyncUtil.QueueAsMacro`. 

For example, this menu button starts a Task that takes a while, and 
upon completion it updates cell B1 on Sheet1 using the C API (via an ExcelReference). 

{% highlight csharp %}
[ExcelCommand(MenuName="Async Test", MenuText="Run Later")](ExcelCommand(MenuName=_Async-Test_,-MenuText=_Run-Later_)) 
public static void RunLater() 
{ 
    Task.Factory.StartNew( () => Thread.Sleep(5000) ) 
    .ContinueWith(t => 
        ExcelAsyncUtil.QueueAsMacro(() => 
        { 
            var refB1 = new ExcelReference(0,0,1,1, "Sheet1"); 
            refB1.SetValue("Done!"); 
        })); 
} 
{% endhighlight %}


Internally this is implemented by adding the QueueAsMacro delegate on a queue, and (normally) posting a WM_SYNCMACRO event.  The delegate is then dequequed by a SyncMacro function that is run as an Excel  xlfregister ed macro.  (There are several functions called "SyncMacro" in the ExcelDNA, the one that is actually registered is in Exceldna.cpp.)

A new NativeWindow on Excel's main event loop traps WM_SYNCMACRO and WM_TIMER events and attempts to run SyncMacro on the main thread.  If the attempt fails then the timer is reset for 250ms later.   A test is performed to ensure SyncMacro is not run while the user is editing a formula.  (The same NativeWindow is also used for RTD processing.)
