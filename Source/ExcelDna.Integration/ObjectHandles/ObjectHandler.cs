﻿using System;
using System.Collections.Concurrent;
using System.Collections.Generic;

namespace ExcelDna.Integration.ObjectHandles
{
    public class ObjectHandler
    {
        private static ConcurrentDictionary<string, HandleInfo> _objects = new ConcurrentDictionary<string, HandleInfo>();

        // Tries to get an existing handle for the given function and parameters.
        // If there is no existing handle, creates a new handle with the target provided by evaluating the delegate 'func'
        // (with the given function and parameters).
        public static object GetHandle(string callerFunctionName, object callerParameters, object userObject)
        {
            object result = ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, () =>
            {
                if (userObject is LazyLambda lazyLambda)
                    userObject = lazyLambda.Invoke();

                var handleInfo = new HandleInfo(callerFunctionName, userObject);
                _objects.TryAdd(handleInfo.Handle, handleInfo);
                return handleInfo;
            });

            return result;
        }

        public static string GetHandle(string tag, object userObject)
        {
            var handleInfo = new HandleInfo(tag, userObject);
            _objects.TryAdd(handleInfo.Handle, handleInfo);
            return handleInfo.Handle;
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

        public static void Remove(string handle)
        {
            HandleInfo value;
            if (_objects.TryRemove(handle, out value))
            {
                var disp = value.UserObject as IDisposable;
                if (disp != null)
                {
                    disp.Dispose();
                }
            }
        }
    }
}
