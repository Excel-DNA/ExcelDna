using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelDna.Integration.ObjectHandles
{
    class HandleInfo : IExcelObservable, IDisposable
    {
        // Global index used in handle names
        static int HandleIndex;

        // This is the information we need to Refresh the Target
        // Never changes
        public readonly string ObjectType;
        public readonly object[] Parameters;
        public readonly ObjectHandler Handler;

        // The Target and Handle can change
        public string Handle;
        public DateTime LastUpdate;

        public readonly object UserObjectNew;

        // Set internally when hooked up to Excel
        public IExcelObserver Observer;

        public HandleInfo(ObjectHandler objectHandler, string objectType, object[] parameters, object userObject)
        {
            // TODO: Complete member initialization
            Handler = objectHandler;
            ObjectType = objectType;
            Parameters = parameters;

            Handle = string.Format("{0}:{1}", objectType, HandleIndex++);
            LastUpdate = DateTime.Now;

            UserObjectNew = userObject;
        }

        // This call is made (once) from Excel to subscribe to the topic.
        public IDisposable Subscribe(IExcelObserver observer)
        {
            // We know this will only be called once, so we take some adventurous shortcuts (like returning 'this')
            Observer = observer;
            Observer.OnNext(Handle);
            return this;
        }

        // Called from the ObjectHandler
        internal void Update()
        {
            Handle = string.Format("{0}:{1}", ObjectType, HandleIndex++); ;
            LastUpdate = DateTime.Now;          // Might be used to decide when or how often to refresh
            if (Observer != null)
                Observer.OnNext(Handle);        // Triggers the update sending the new handle to Excel
        }

        public void Dispose()
        {
            Handler.Remove(this);               // Called when last instance of this topic is removed from the current session
        }
    }

    class ObjectHandler
    {
        Dictionary<string, HandleInfo> _objects = new Dictionary<string, HandleInfo>();

        public ObjectHandler()
        {
        }

        // Tries to get an existing handle for the given object type and parameters.
        // If there is no existing handle, creates a new handle with the target provided by evaluating the delegate 'func'
        // (with the given object type and parameters).
        public object GetHandleNew(string callerFunctionName, object callerParameters, object userObject)
        {
            return ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, () =>
            {
                //var target = _dataService.ProcessRequest(objectType, parameters);
                var handleInfo = new HandleInfo(this, callerFunctionName, null, userObject);
                _objects.Add(handleInfo.Handle, handleInfo);
                return handleInfo;
            });
        }


        public bool TryGetObjectNew(string handle, out object value)
        {
            HandleInfo handleInfo;
            if (_objects.TryGetValue(handle, out handleInfo))
            {
                value = handleInfo.UserObjectNew;
                return true;
            }
            value = null;
            return false;
        }

        internal void Remove(HandleInfo handleInfo)
        {
            object value;
            if (TryGetObjectNew(handleInfo.Handle, out value))
            {
                _objects.Remove(handleInfo.Handle);
                var disp = value as IDisposable;
                if (disp != null)
                {
                    disp.Dispose();
                }
            }
        }
    }
}
