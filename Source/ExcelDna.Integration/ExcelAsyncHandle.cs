/*
  Copyright (C) 2005-2013 Govert van Drimmelen

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
using System.Diagnostics;
using ExcelDna.Integration.Rtd;

namespace ExcelDna.Integration
{
    // NOTE: The types and methods in this file are called via reflection from ExcelDna.Loader - IntegrationMarshalHelpers.cs

    // CONSIDER: Should this rather be an interface? Does it matter?
    public abstract class ExcelAsyncHandle
    {
        public abstract bool SetResult(object result);
        public abstract bool SetException(Exception exception);
    }

    internal class ExcelAsyncHandleNative : ExcelAsyncHandle
    {
        // NOTE: This field is read by reflection from ExcelDna.Loader.IntegrationMarshalHelpers
        readonly IntPtr _handle;

        // NOTE: This constructor is accessed by reflection from ExcelDna.Loader.IntegrationMarshalHelpers
        ExcelAsyncHandleNative(IntPtr handle)
        {
            _handle = handle;
        }

        public override bool SetResult(object result)
        {
            // Typically called from a completely independent thread, e.g. a threadpool worker,
            // so any exception here would crash Excel.
            object unusedResult;
            XlCall.XlReturn callReturn = XlCall.TryExcel(XlCall.xlAsyncReturn, out unusedResult, this, result);
            if (callReturn == XlCall.XlReturn.XlReturnSuccess)
            {
                // The normal case - value has been accepted
                return true;
            }

            if (callReturn == XlCall.XlReturn.XlReturnInvAsynchronousContext)
            {
                // This is expected sometimes (e.g. calculation was cancelled)
                // Excel will show #VALUE
                Debug.WriteLine("Warning: InvalidAsyncContext returned from xlAsyncReturn");
                return false;
            }

            // This is never unexpected
            Debug.WriteLine("Error: Unexpected error from xlAsyncReturn");
            return false;
        }

        // Calls the Excel-DNA UnhandledExceptionHandler (which by default returns #VALUE to Excel).
        public override bool SetException(Exception exception)
        {
            object result = ExcelIntegration.HandleUnhandledException(exception);
            return SetResult(result);
        }

        // We could do something like this:
        //public static bool SetResults(object[] asyncHandles, object[] results)
        //{
        //XlCall.XlReturn callReturn = XlCall.TryExcel(XlCall.xlAsyncReturn, out unusedResult, asyncHandles, results);
        //}
    }

    internal class ExcelAsyncHandleObservable : ExcelAsyncHandle, IExcelObservable
    {
        bool _hasResult;
        object _result;
        Exception _exception;
        IExcelObserver _observer;
        readonly object _observerLock = new object();

        public override bool SetResult(object result)
        {
            lock (_observerLock)
            {
                if (_hasResult) throw new InvalidOperationException("ExcelAsyncHandle Result can be set only once.");

                _hasResult = true;
                _result = result;

                if (_observer != null)
                {
                    _observer.OnNext(result);
                    _observer.OnCompleted();
                }
                return true;
            }
        }

        public override bool SetException(Exception exception)
        {
            lock (_observerLock)
            {
                if (_hasResult) throw new InvalidOperationException("ExcelAsyncHandle Result can be set only once.");

                _hasResult = true;
                _exception = exception;

                if (_observer != null)
                {
                    _observer.OnError(exception);
                }
                return true;
            }
        }

        public IDisposable Subscribe(IExcelObserver observer)
        {
            lock (_observerLock)
            {
                if (_observer != null) throw new InvalidOperationException("Only single Subscription allowed.");
                _observer = observer;

                if (_hasResult)
                {
                    if (_exception != null)
                    {
                        _observer.OnError(_exception);
                    }
                    else
                    {
                        _observer.OnNext(_result);
                        _observer.OnCompleted();
                    }
                }
            }

            return DummyDisposable.Instance;
        }

        class DummyDisposable : IDisposable
        {
            public static readonly DummyDisposable Instance = new DummyDisposable();

            private DummyDisposable()
            {
            }

            public void Dispose()
            {
            }
        }

    }

}
