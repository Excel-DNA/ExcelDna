using System;
using System.Diagnostics;
using System.Reactive.Threading.Tasks;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelDna.Integration.RxExcel
{
    public static class RxExcel
    {
        public static IExcelObservable ToExcelObservable<T>(this IObservable<T> observable)
        {
            return new ExcelObservable<T>(observable);
        }

        public static object Observe<T>(string callerFunctionName, object callerParameters, Func<IObservable<T>> observableSource)
        {
            return ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, () => observableSource().ToExcelObservable());
        }

        public static object Observe<T>(string callerFunctionName, object callerParameters, Func<string, object, IObservable<T>> observableSource)
        {
            return ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, () => observableSource(callerFunctionName, callerParameters).ToExcelObservable());
        }

        // TODO: Tasks may be used with Excel-DNA async without using Rx.
        public static object Observe<T>(string callerFunctionName, object callerParameters, Func<Task<T>> taskSource)
        {
            return Observe(callerFunctionName, callerParameters, () => taskSource().ToObservable());
        }

        public static object Observe<T>(string callerFunctionName, object callerParameters, Func<string, object, Task<T>> taskSource)
        {
            return Observe(callerFunctionName, callerParameters, () => taskSource(callerFunctionName, callerParameters).ToObservable());
        }
    }

    public class ExcelObservable<T> : IExcelObservable
    {
        readonly IObservable<T> _observable;

        public ExcelObservable(IObservable<T> observable)
        {
            _observable = observable;
        }

        public IDisposable Subscribe(IExcelObserver observer)
        {
#if DEBUG
            return new DebuggingDisposable(_observable.Subscribe(value => observer.OnNext(value), observer.OnError, observer.OnCompleted));
#else
            return _observable.Subscribe(value => observer.OnNext(value), observer.OnError, observer.OnCompleted);
#endif
        }
    }

    public class DebuggingDisposable : IDisposable
    {
        readonly IDisposable _disposable;

        public DebuggingDisposable(IDisposable disposable)
        {
            _disposable = disposable;
        }

        public void Dispose()
        {
            Debug.Print("Disposing...");
            _disposable.Dispose();
        }
    }

    // TODO: Use for Tasks -> Excel-DNA async directly.
    /// <summary>
    /// Represents an IDisposable that can be checked for cancellation status.
    /// </summary>
    public sealed class CancellationDisposable : IDisposable
    {
        CancellationTokenSource cts;

        /// <summary>
        /// Constructs a new CancellationDisposable that uses an existing CancellationTokenSource.
        /// </summary>
        public CancellationDisposable(CancellationTokenSource cts)
        {
            if (cts == null)
                throw new ArgumentNullException("cts");

            this.cts = cts;
        }

        /// <summary>
        /// Constructs a new CancellationDisposable that uses a new CancellationTokenSource.
        /// </summary>
        public CancellationDisposable()
            : this(new CancellationTokenSource())
        {
        }

        /// <summary>
        /// Gets the CancellationToken used by this CancellationDisposable.
        /// </summary>
        public CancellationToken Token { get { return cts.Token; } }

        /// <summary>
        /// Cancels the CancellationTokenSource.
        /// </summary>
        public void Dispose()
        {
            cts.Cancel();
        }
    }
}
