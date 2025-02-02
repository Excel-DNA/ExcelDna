using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelDna.Registration.Utils
{
    // Helpers for creating and using Task-based functions with Excel 2010+ native async support
    public static class NativeAsyncTaskUtil
    {
        // CONSIDER: Is it important to call this, or will Excel / Excel-DNA clean up when if the add-in gets unloaded?
        public static void Uninitialize()
        {
            Integration.NativeAsyncTaskUtil.Uninitialize();
        }

        // We could have implemented this by just taking a Task<T>,
        // but keep the Func<Task<T>> indirection to be consistent with the similar calls in AsyncTaskUtil.
        public static void RunTask<TResult>(Func<Task<TResult>> taskSource, ExcelAsyncHandle asyncHandle)
        {
            Integration.NativeAsyncTaskUtil.RunTask(taskSource, asyncHandle);
        }

        public static void RunTaskWithCancellation<TResult>(Func<CancellationToken, Task<TResult>> taskSource, ExcelAsyncHandle asyncHandle)
        {
            Integration.NativeAsyncTaskUtil.RunTaskWithCancellation(taskSource, asyncHandle);
        }

        public static void RunAsTask<TResult>(Func<TResult> function, ExcelAsyncHandle asyncHandle)
        {
            Integration.NativeAsyncTaskUtil.RunAsTask(function, asyncHandle);
        }

        public static void RunAsTaskWithCancellation<TResult>(Func<CancellationToken, TResult> function, ExcelAsyncHandle asyncHandle)
        {
            Integration.NativeAsyncTaskUtil.RunAsTaskWithCancellation(function, asyncHandle);
        }
    }
}
