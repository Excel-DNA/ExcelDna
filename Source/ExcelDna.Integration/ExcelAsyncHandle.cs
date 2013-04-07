using System;
using System.Diagnostics;

namespace ExcelDna.Integration
{
    // Class or struct?
    public class ExcelAsyncHandle
    {
        // NOTE: This field is read by reflection from ExcelDna.Loader.IntegrationMarshalHelpers
        readonly IntPtr _handle;

        // NOTE: This constructor is accessed by reflection from ExcelDna.Loader.IntegrationMarshalHelpers
        ExcelAsyncHandle(IntPtr handle)
        {
            _handle = handle;
        }

        public bool SetResult(object result)
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
        public bool SetException(Exception exception)
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
}
