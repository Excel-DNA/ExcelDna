using System;
using ExcelDna.Integration;

namespace ExcelDna.Registration.Utils
{
    public static class ObservableRtdUtil
    {
        public static object Observe<T>(string callerFunctionName, object callerParameters, Func<IObservable<T>> observableSource)
        {
            return ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, observableSource);
        }
    }
}
