using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelDna.AsyncSample
{
    internal static class ExcelTaskUtil
    {
        // Careful - this might only work as long as the task is not shared between calls, since cancellation cancels that task
        // Another implementation via Reactive Extension is Task.ToObservable() (in System.Reactive.Linq.dll) with RxExcel
        public static IExcelObservable ToExcelObservable<TResult>(this Task<TResult> task)
        {
            if (task == null)
            {
                throw new ArgumentNullException("task");
            }

            return new ExcelTaskObservable<TResult>(task);
        }

        public static IExcelObservable ToExcelObservable<TResult>(this Task<TResult> task, CancellationTokenSource cts)
        {
            if (task == null)
            {
                throw new ArgumentNullException("task");
            }

            return new ExcelTaskObservable<TResult>(task, cts);
        }

        public static object RunTask<TResult>(string callerFunctionName, object callerParameters, Func<CancellationToken, Task<TResult>> taskSource)
        {
            return ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, delegate
                {
                    var cts = new CancellationTokenSource();
                    var task = taskSource(cts.Token);
                    return new ExcelTaskObservable<TResult>(task, cts);
                });
        }

        public static object RunTask<TResult>(string callerFunctionName, object callerParameters, Func<Task<TResult>> taskSource)
        {
            return ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, delegate
            {
                var task = taskSource();
                return new ExcelTaskObservable<TResult>(task);
            });
        }

	    public static object RunAsTask<TResult>(string callerFunctionName, object callerParameters, Func<CancellationToken, TResult> function)
	    {
            return RunTask(callerFunctionName, callerParameters, cancellationToken => Task.Factory.StartNew<TResult>(() => function(cancellationToken), cancellationToken));
	    }

	    public static object RunAsTask<TResult>(string callerFunctionName, object callerParameters, Func<TResult> function)
	    {
		    return RunTask(callerFunctionName, callerParameters, () => Task.Factory.StartNew<TResult>(function));
	    }

        // Helper class to wrap a Task in an Observable - allowing one Subscriber.
        class ExcelTaskObservable<TResult> : IExcelObservable
        {
            readonly Task<TResult> _task;
            readonly CancellationTokenSource _cts;

            public ExcelTaskObservable(Task<TResult> task)
            {
                _task = task;
            }

            public ExcelTaskObservable(Task<TResult> task, CancellationTokenSource cts)
                : this(task)
            {
                _cts = cts;
            }

            public IDisposable Subscribe(IExcelObserver observer)
            {
                switch (_task.Status)
                {
                    case TaskStatus.RanToCompletion:
                        observer.OnNext(_task.Result);
                        observer.OnCompleted();
                        break;
                    case TaskStatus.Faulted:
                        observer.OnError(_task.Exception.InnerException);
                        break;
                    case TaskStatus.Canceled:
                        observer.OnError(new TaskCanceledException(_task));
                        break;
                    default:
                        _task.ContinueWith(t =>
                        {
                            switch (t.Status)
                            {
                                case TaskStatus.RanToCompletion:
                                    observer.OnNext(t.Result);
                                    observer.OnCompleted();
                                    break;
                                case TaskStatus.Faulted:
                                    observer.OnError(t.Exception.InnerException);
                                    break;
                                case TaskStatus.Canceled:
                                    observer.OnError(new TaskCanceledException(t));
                                    break;
                            }
                        });
                        break;
                }

                // Check for cancellation support
                if (_cts != null)
                {
                    return new CancellationDisposable(_cts);
                }
                // No cancellation
                return DefaultDisposable.Instance;
            }
        }

        sealed class DefaultDisposable : IDisposable
        {

            public static readonly DefaultDisposable Instance = new DefaultDisposable();
            
            // Prevent external instantiation
            DefaultDisposable()
            {
            }

            public void Dispose()
            {
                // no op
            }
        }

        sealed class CancellationDisposable : IDisposable
        {

            readonly CancellationTokenSource cts;
            public CancellationDisposable(CancellationTokenSource cts)
            {
                if (cts == null)
                {
                    throw new ArgumentNullException("cts");
                }

                this.cts = cts;
            }

            public CancellationDisposable()
                : this(new CancellationTokenSource())
            {
            }

            public CancellationToken Token
            {
                get { return cts.Token; }
            }

            public void Dispose()
            {
                cts.Cancel();
            }
        }
    }
}