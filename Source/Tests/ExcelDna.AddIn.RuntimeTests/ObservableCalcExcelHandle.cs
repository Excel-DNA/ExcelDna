using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDna.AddIn.RuntimeTests
{
    internal class ObservableCalcExcelHandle : IObservable<CalcExcelHandle>
    {
        private CalcExcelHandle c;
        private List<IObserver<CalcExcelHandle>> observers;

        public ObservableCalcExcelHandle(double d1, double d2)
        {
            this.c = new CalcExcelHandle(d1, d2);
            observers = new List<IObserver<CalcExcelHandle>>();
        }

        public IDisposable Subscribe(IObserver<CalcExcelHandle> observer)
        {
            observers.Add(observer);
            observer.OnNext(c);
            return new ActionDisposable(() => observers.Remove(observer));
        }

        private class ActionDisposable : IDisposable
        {
            private Action disposeAction;

            public ActionDisposable(Action disposeAction)
            {
                this.disposeAction = disposeAction;
            }

            public void Dispose()
            {
                disposeAction();
            }
        }
    }
}
