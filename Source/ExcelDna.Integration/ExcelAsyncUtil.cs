//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Diagnostics;
using System.Threading;
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
    public delegate void ExcelFuncAsyncHandle(ExcelAsyncHandle handle);
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
        [Obsolete("ExcelAsyncUtil.Initialize is no longer required. The call can be removed.")]
        public static void Initialize() {}
        [Obsolete("ExcelAsyncUtil.Uninitialize is no longer required. The call can be removed.")]
        public static void Uninitialize() {}

        // Async observable support
        // This is the most general RTD registration
        public static object Observe(string callerFunctionName, object callerParameters, ExcelObservableSource observableSource)
        {
            return AsyncObservableImpl.ProcessObservable(callerFunctionName, callerParameters, observableSource);
        }

        // Async function support
        public static object Run(string callerFunctionName, object callerParameters, ExcelFunc asyncFunc)
        {
            Debug.Print("ExcelAsyncUtil.Run - {0} : {1}", callerFunctionName, callerParameters);
            return AsyncObservableImpl.ProcessFunc(callerFunctionName, callerParameters, asyncFunc);
        }

        // Async function with ExcelAsyncHandle
        // The function will run on the main thread (like an Excel 2010+ native async function), 
        // but can spawn a thread and return the value later.
        public static object Run(string callerFunctionName, object callerParameters, ExcelFuncAsyncHandle asyncFunc)
        {
            return AsyncObservableImpl.ProcessFuncAsyncHandle(callerFunctionName, callerParameters, asyncFunc);
        }

        // Async macro support
        public static void QueueMacro(string macroName)
        {
            QueueAsMacro(RunMacro, macroName);
        }

        public static void QueueAsMacro(ExcelAction action)
        {
            SendOrPostCallback callback = delegate { action(); };
            QueueAsMacro(callback, null);
        }

        public static void QueueAsMacro(SendOrPostCallback callback, object state)
        {
            SynchronizationManager.RunMacroSynchronization.RunAsMacroAsync(callback, state);
        }

        static void RunMacro(object macroName)
        {
            XlCall.Excel(XlCall.xlcRun, macroName);
        }

        #region Async calculation events
        // CONSIDER: Do we need to unregister these when unloaded / reloaded?

        static ExcelAction _calculationCanceled = null;
        public static event ExcelAction CalculationCanceled
        {
            add
            {
                if (_calculationCanceled != null)
                {
                    // We've set it up already
                    // just add the delegate to the event
                    _calculationCanceled += value;
                }
                else
                {
                    // First time - register event handler
                    _calculationCanceled = value;
                    double result = (double)XlCall.Excel(XlCall.xlEventRegister, "CalculationCanceled", XlCall.xleventCalculationCanceled);
                    if (result == 0)
                    {
                        // CONSIDER: Is there a better way to handle this unexpected error?
                        throw new XlCallException(XlCall.XlReturn.XlReturnFailed);
                    }
                }
            }
            remove
            {
                _calculationCanceled -= value;
                if (_calculationCanceled == null)
                {
                    XlCall.Excel(XlCall.xlEventRegister, null, XlCall.xleventCalculationCanceled);
                }
            }
        }

        internal static void OnCalculationCanceled()
        {
            if (_calculationCanceled != null) _calculationCanceled();
        }

        static ExcelAction _calculationEnded = null;
        public static event ExcelAction CalculationEnded
        {
            add
            {
                if (_calculationEnded != null)
                {
                    // We've set it up already
                    // just add the delegate to the event
                    _calculationEnded += value;
                }
                else
                {
                    // First time - register event handler
                    _calculationEnded = value;
                    double result = (double)XlCall.Excel(XlCall.xlEventRegister, "CalculationEnded", XlCall.xleventCalculationEnded);
                    if (result == 0)
                    {
                        // CONSIDER: Is there a better way to handle this unexpected error?
                        throw new XlCallException(XlCall.XlReturn.XlReturnFailed);
                    }
                }
            }
            remove
            {
                _calculationEnded -= value;
                if (_calculationEnded == null)
                {
                    XlCall.Excel(XlCall.xlEventRegister, null, XlCall.xleventCalculationEnded);
                }
            }
        }

        internal static void OnCalculationEnded()
        {
            if (_calculationEnded != null) _calculationEnded();
        }
        #endregion
    }
}