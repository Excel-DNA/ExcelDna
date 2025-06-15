namespace ExcelDna.AddIn.RuntimeTests
{
    internal class ObservableString : IObservable<string>
    {
        private string s;
        private List<IObserver<string>> observers;

        public ObservableString(string s)
        {
            this.s = s;
            observers = new List<IObserver<string>>();
        }

        public IDisposable Subscribe(IObserver<string> observer)
        {
            observers.Add(observer);
            observer.OnNext(s);
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
