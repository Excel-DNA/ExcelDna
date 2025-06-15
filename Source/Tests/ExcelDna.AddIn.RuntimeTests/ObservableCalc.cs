using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDna.AddIn.RuntimeTests
{
    internal class ObservableCalc : IObservable<Calc>
    {
        private Calc c;
        private List<IObserver<Calc>> observers;

        public ObservableCalc(double d1, double d2)
        {
            this.c = new Calc(d1, d2);
            observers = new List<IObserver<Calc>>();
        }

        public IDisposable Subscribe(IObserver<Calc> observer)
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
