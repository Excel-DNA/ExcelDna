/*
  Copyright (C) 2005-2012 Govert van Drimmelen

  This software is provided 'as-is', without any express or implied
  warranty.  In no event will the authors be held liable for any damages
  arising from the use of this software.

  Permission is granted to anyone to use this software for any purpose,
  including commercial applications, and to alter it and redistribute it
  freely, subject to the following restrictions:

  1. The origin of this software must not be misrepresented; you must not
     claim that you wrote the original software. If you use this software
     in a product, an acknowledgment in the product documentation would be
     appreciated but is not required.
  2. Altered source versions must be plainly marked as such, and must not be
     misrepresented as being the original software.
  3. This notice may not be removed or altered from any source distribution.


  Govert van Drimmelen
  govert@icon.co.za
*/

using System;
using System.Threading;
using ExcelDna.ComInterop;
using ExcelDna.Integration.Rtd;

namespace ExcelDna.Integration
{
    // Introduction to Rx: http://www.introtorx.com/

    // Task.ToObservable: http://blogs.msdn.com/b/pfxteam/archive/2010/04/04/9990349.aspx

    // Pattern for making an Observable: http://msdn.microsoft.com/en-us/library/dd990377.aspx
    // Task.ToObservable: http://blogs.msdn.com/b/pfxteam/archive/2010/04/04/9990349.aspx

    // Action and Func are not defined in .NET 2.0
    public delegate void ExcelAction();
    public delegate object ExcelFunc();
    public delegate IExcelObservable ExcelObservableSource();

    // The next two interfaces would be regular System.IObservable<object> if we could target .NET 4.
    // TODO: Make an adapter to make it easy to integrate with .NET 4.
    public interface IExcelObservable
    {
        IDisposable Subscribe(IExcelObserver observer);
    }

    public interface IExcelObserver
    {
        void OnCompleted();
        void OnError(Exception exception);
        void OnNext(object value);
    }

    public static class ExcelAsyncUtil
    {
        // Initialization - must be called from a macro context (e.g. AutoOpen)
        // For now only installs the syncmanager.
        public static void Initialize()
        {
            SynchronizationManager.Install();
        }

        public static void Uninitialize()
        {
            SynchronizationManager.Uninstall();
        }

        // Async observable support
        // This is the most general RTD registration
        // TODO: This should not be called from a ThreadSafe function. Check...?
        public static object Observe(string callerFunctionName, object callerParameters, ExcelObservableSource observableSource)
        {
            return AsyncObservableImpl.ProcessObservable(callerFunctionName, callerParameters, observableSource);
        }

        // Async function support
        public static object Run(string callerFunctionName, object callerParameters, ExcelFunc asyncFunc)
        {
            return AsyncObservableImpl.ProcessFunc(callerFunctionName, callerParameters, asyncFunc);
        }

        // Async macro support
        public static void QueueMacro(string macroName)
        {
            QueueAsMacro(RunMacro, macroName);
        }

        public static void QueueAsMacro(ExcelAction action)
        {
            QueueAsMacro(delegate { action(); }, null);
        }

        public static void QueueAsMacro(SendOrPostCallback callback, object state)
        {
            if (!SynchronizationManager.IsInstalled)
                throw new InvalidOperationException("SynchronizationManager is not registered.");

            SynchronizationManager.RunMacroSynchronization.RunAsMacroAsync(callback, state);
        }

        static void RunMacro(object macroName)
        {
            XlCall.Excel(XlCall.xlcRun, macroName);
        }
    }
}