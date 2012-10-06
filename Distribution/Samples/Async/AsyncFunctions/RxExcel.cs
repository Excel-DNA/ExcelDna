using System;
using System.Diagnostics;
using System.Reactive.Threading.Tasks;
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
            return new DebuggingDisposable(_observable.Subscribe(value => observer.OnNext(value), observer.OnError, observer.OnCompleted));
            //return _observable.Subscribe(value => observer.OnNext(value), observer.OnError, observer.OnCompleted);
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
            Debug.Print("Disposing!");
            _disposable.Dispose();
        }
    }
}
