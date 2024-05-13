using System;

namespace ExcelDna.Integration.ObjectHandles
{
    internal class HandleInfo : IExcelObservable, IDisposable
    {
        // Global index used in handle names
        private static long HandleIndex;

        private readonly ObjectHandler Handler;

        public readonly string Handle;
        public readonly object UserObject;

        public HandleInfo(ObjectHandler objectHandler, string objectType, object userObject)
        {
            Handler = objectHandler;
            Handle = string.Format("{0}:{1}", objectType, HandleIndex++);
            UserObject = userObject;
        }

        // This call is made (once) from Excel to subscribe to the topic.
        public IDisposable Subscribe(IExcelObserver observer)
        {
            // We know this will only be called once, so we take some adventurous shortcuts (like returning 'this')
            observer.OnNext(Handle);
            return this;
        }

        public void Dispose()
        {
            Handler.Remove(this);               // Called when last instance of this topic is removed from the current session
        }
    }
}
