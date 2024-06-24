using System;

namespace ExcelDna.Integration
{
    public class ExcelObjectHandle<T> : IDisposable
    {
        private bool disposedValue;

        public T Object { get; }

        public ExcelObjectHandle(T o)
        {
            Object = o;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    (Object as IDisposable)?.Dispose();
                }

                disposedValue = true;
            }
        }

        ~ExcelObjectHandle()
        {
            Dispose(disposing: false);
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
