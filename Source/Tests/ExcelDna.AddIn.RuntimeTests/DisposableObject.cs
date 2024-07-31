namespace ExcelDna.AddIn.RuntimeTests
{
    public class DisposableObject : IDisposable
    {
        public static int ObjectsCount { get; private set; } = 0;
        public static int CreatedObjectsCount { get; private set; } = 0;
        private bool disposedValue;

        public DisposableObject()
        {
            ++CreatedObjectsCount;
            ++ObjectsCount;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    --ObjectsCount;
                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
