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
        public IHasRowVersion Target;
        public string Handle;
        public DateTime LastUpdate;

        public readonly ExcelObjectHandle UserObjectNew;

        // Set internally when hooked up to Excel
        public IExcelObserver Observer;

        public HandleInfo(ObjectHandler objectHandler, string objectType, object[] parameters, ExcelObjectHandle userObject, IHasRowVersion target)
        {
            // TODO: Complete member initialization
            Handler = objectHandler;
            ObjectType = objectType;
            Parameters = parameters;

            Target = target;
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
        internal void Update(IHasRowVersion target)
        {
            Target = target;
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
        IDataService _dataService;

        public ObjectHandler(IDataService dataService)
        {
            _dataService = dataService;
        }

        // Tries to get an existing handle for the given object type and parameters.
        // If there is no existing handle, creates a new handle with the target provided by evaluating the delegate 'func'
        // (with the given object type and parameters).
        public object GetHandleNew(string objectType, ExcelObjectHandle userObject)
        {
            return ExcelAsyncUtil.Observe(objectType, userObject.CallerParameters, () =>
            {
                //var target = _dataService.ProcessRequest(objectType, parameters);
                var handleInfo = new HandleInfo(this, objectType, userObject.CallerParameters, userObject, null);
                _objects.Add(handleInfo.Handle, handleInfo);
                return handleInfo;
            });
        }

        public bool TryGetObject(string handle, out object value)
        {
            HandleInfo handleInfo;
            if (_objects.TryGetValue(handle, out handleInfo))
            {
                value = handleInfo.Target;
                return true;
            }
            value = null;
            return false;
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
            if (TryGetObject(handleInfo.Handle, out value))
            {
                _objects.Remove(handleInfo.Handle);
                var disp = value as IDisposable;
                if (disp != null)
                {
                    disp.Dispose();
                }
            }
        }

        // Forces a Refresh for all handles, not checking the RowVersions
        public void RefreshAll()
        {
            // We make a copy of the active HandleInfos (so that we can update the dictionary itself in the loop)
            var activeHandles = _objects.Values.ToArray();
            foreach (var handleInfo in activeHandles)
            {
                Refresh(handleInfo);
            }
        }

        // Forces a Refresh for a particular handle, not checking the RowVersion
        public void Refresh(HandleInfo handleInfo)
        {
            _objects.Remove(handleInfo.Handle);
            var target = _dataService.ProcessRequest(handleInfo.ObjectType, handleInfo.Parameters);
            handleInfo.Update(target);
            _objects.Add(handleInfo.Handle, handleInfo);
        }

        // Calls the data service with the existing rowversions, and does a selective update
        // Assumes that the DataService returns nulls for objects not updated
        // Could be changed to check the rowversions, or some other way of looking it up
        public void UpdateAll()
        {
            var requestInfos = new List<Tuple<string, object[], ulong>>();
            var activeHandles = _objects.Values.ToArray();
            foreach (var handleInfo in activeHandles)
            {
                requestInfos.Add(Tuple.Create(handleInfo.ObjectType, handleInfo.Parameters, handleInfo.Target.RowVersion));
            }

            var updates = _dataService.ProcessUpdateRequests(requestInfos);
            for (int i = 0; i < activeHandles.Length; i++)
            {
                var handleInfo = activeHandles[i];
                var target = updates[i];
                if (target != null)
                {
                    _objects.Remove(handleInfo.Handle);
                    handleInfo.Update(target);
                    _objects.Add(handleInfo.Handle, handleInfo);
                }
            }
        }
    }
}
