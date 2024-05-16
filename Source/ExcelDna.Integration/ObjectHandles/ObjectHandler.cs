using System;
using System.Collections.Generic;

namespace ExcelDna.Integration.ObjectHandles
{
    internal class ObjectHandler
    {
        private static Dictionary<string, HandleInfo> _objects = new Dictionary<string, HandleInfo>();

        // Tries to get an existing handle for the given function and parameters.
        // If there is no existing handle, creates a new handle with the target provided by evaluating the delegate 'func'
        // (with the given function and parameters).
        public static object GetHandle(string callerFunctionName, object callerParameters, object userObject, out bool newHandle)
        {
            bool newHandleCreated = false;
            object result = ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, () =>
            {
                var handleInfo = new HandleInfo(callerFunctionName, userObject);
                _objects.Add(handleInfo.Handle, handleInfo);
                newHandleCreated = true;
                return handleInfo;
            });
            newHandle = newHandleCreated;

            return result;
        }

        public static bool TryGetObject(string handle, out object value)
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

        public static void Remove(HandleInfo handleInfo)
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
    }
}
