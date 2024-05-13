using System;
using System.Collections.Generic;

namespace ExcelDna.Integration.ObjectHandles
{
    class ObjectHandler
    {
        Dictionary<string, HandleInfo> _objects = new Dictionary<string, HandleInfo>();

        public ObjectHandler()
        {
        }

        // Tries to get an existing handle for the given object type and parameters.
        // If there is no existing handle, creates a new handle with the target provided by evaluating the delegate 'func'
        // (with the given object type and parameters).
        public object GetHandleNew(string callerFunctionName, object callerParameters, object userObject, out bool newHandle)
        {
            bool newHandleCreated = false;
            object result = ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, () =>
            {
                //var target = _dataService.ProcessRequest(objectType, parameters);
                var handleInfo = new HandleInfo(this, callerFunctionName, userObject);
                _objects.Add(handleInfo.Handle, handleInfo);
                newHandleCreated = true;
                return handleInfo;
            });
            newHandle = newHandleCreated;
            return result;
        }

        public bool TryGetObjectNew(string handle, out object value)
        {
            HandleInfo handleInfo;
            if (_objects.TryGetValue(handle, out handleInfo))
            {
                value = handleInfo.UserObject;
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
