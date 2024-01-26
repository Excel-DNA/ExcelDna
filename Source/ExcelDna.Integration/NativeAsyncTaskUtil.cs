using System;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelDna.Integration
{
    // Helpers for creating and using Task-based functions with Excel 2010+ native async support
    public static class NativeAsyncTaskUtil
    {
        // Everything here has to happen on the main thead.
        static int _initializeCount;

        public static void Initialize()
        {
            if (_initializeCount == 0)
            {
                ExcelAsyncUtil.CalculationCanceled += CalculationCanceled;
                ExcelAsyncUtil.CalculationEnded += CalculationEnded;
            }
            _initializeCount++;
        }

        // CONSIDER: Is it important to call this, or will Excel / Excel-DNA clean up when if the add-in gets unloaded?
        public static void Uninitialize()
        {
            _initializeCount--;
            if (_initializeCount == 0)
            {
                ExcelAsyncUtil.CalculationCanceled -= CalculationCanceled;
                ExcelAsyncUtil.CalculationEnded -= CalculationEnded;
            }
        }

        // Cancellation support
        // We keep a CancellationTokenSource around, and set to a new one whenever a calculation has finished.
        static CancellationTokenSource _cancellation = new CancellationTokenSource();
        static void CalculationCanceled()
        {
            _cancellation.Cancel();
        }

        static void CalculationEnded()
        {
            // Set to a fresh CancellationTokenSource after every calculation is complete
            _cancellation.Dispose();
            _cancellation = new CancellationTokenSource();
        }

        // We could have implemented this by just taking a Task<T>,
        // but keep the Func<Task<T>> indirection to be consistent with the similar calls in AsyncTaskUtil.
        public static void RunTask<TResult>(Func<Task<TResult>> taskSource, ExcelAsyncHandle asyncHandle)
        {
            var task = taskSource();
            task.ContinueWith(t =>
            {
                try
                {
                    // task.Result will throw an AggregateException if there was an error
                    asyncHandle.SetResult(t.Result);
                }
                catch (AggregateException ex)
                {
                    // There may be multiple exceptions...
                    // Do we have to call Handle?
                    asyncHandle.SetException(ex.InnerException);
                }

                // Unhandled exceptions here will crash Excel 
                // and leave open workbooks in an unrecoverable state...

            }, TaskContinuationOptions.NotOnCanceled);
        }

        public static void RunTaskWithCancellation<TResult>(Func<CancellationToken, Task<TResult>> taskSource, ExcelAsyncHandle asyncHandle)
        {
            var task = taskSource(_cancellation.Token);
            task.ContinueWith(t =>
            {
                try
                {
                    // task.Result will throw an AggregateException if there was an error
                    asyncHandle.SetResult(t.Result);
                }
                catch (AggregateException ex)
                {
                    // There may be multiple exceptions...
                    // Do we have to call Handle?
                    asyncHandle.SetException(ex.InnerException);
                }

                // Unhandled exceptions here will crash Excel 
                // and leave open workbooks in an unrecoverable state...

            }, TaskContinuationOptions.NotOnCanceled);
        }

        public static void RunAsTask<TResult>(Func<TResult> function, ExcelAsyncHandle asyncHandle)
        {
            var task = Task.Factory.StartNew(function);
            task.ContinueWith(t =>
            {
                try
                {
                    // task.Result will throw an AggregateException if there was an error
                    asyncHandle.SetResult(t.Result);
                }
                catch (AggregateException ex)
                {
                    // There may be multiple exceptions...
                    // Do we have to call Handle?
                    asyncHandle.SetException(ex.InnerException);
                }

                // Unhandled exceptions here will crash Excel 
                // and leave open workbooks in an unrecoverable state...

            }, TaskContinuationOptions.NotOnCanceled);
        }

        public static void RunAsTaskWithCancellation<TResult>(Func<CancellationToken, TResult> function, ExcelAsyncHandle asyncHandle)
        {
            CancellationToken cancellationToken = _cancellation.Token;
            var task = Task.Factory.StartNew(() => function(cancellationToken), cancellationToken);
            task.ContinueWith(t =>
            {
                try
                {
                    // task.Result will throw an AggregateException if there was an error
                    asyncHandle.SetResult(task.Result);
                }
                catch (AggregateException ex)
                {
                    // There may be multiple exceptions...
                    // Do we have to call Handle?
                    asyncHandle.SetException(ex.InnerException);
                }

                // Unhandled exceptions here will crash Excel 
                // and leave open workbooks in an unrecoverable state...

            }, TaskContinuationOptions.NotOnCanceled);
        }
    }
}
