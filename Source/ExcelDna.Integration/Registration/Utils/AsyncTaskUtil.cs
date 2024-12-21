using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelDna.Registration.Utils
{
    // Helpers for creating and using Task-based functions with Excel-DNA's RTD-based IObservable support
    public static class AsyncTaskUtil
    {
        public static object RunTask<TResult>(string callerFunctionName, object callerParameters, Func<Task<TResult>> taskSource)
        {
            return ExcelAsyncUtil.RunTask(callerFunctionName, callerParameters, taskSource);
        }

        // Careful - this might only work as long as the task is not shared between calls, since cancellation cancels that task
        public static object RunTaskWithCancellation<TResult>(string callerFunctionName, object callerParameters, Func<CancellationToken, Task<TResult>> taskSource)
        {
            return ExcelAsyncUtil.RunTaskWithCancellation(callerFunctionName, callerParameters, taskSource);
        }

        public static object RunAsTask<TResult>(string callerFunctionName, object callerParameters, Func<TResult> function)
        {
            return ExcelAsyncUtil.RunAsTask(callerFunctionName, callerParameters, function);
        }

        public static object RunAsTaskWithCancellation<TResult>(string callerFunctionName, object callerParameters, Func<CancellationToken, TResult> function)
        {
            return ExcelAsyncUtil.RunAsTaskWithCancellation(callerFunctionName, callerParameters, function);
        }
    }
}
