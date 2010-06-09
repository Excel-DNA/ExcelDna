using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;
using ExcelDna.Integration;
using ExcelDna.ComInterop;
using ExcelDna.ComInterop.ComRegistration;

using HRESULT = System.Int32;
using IID = System.Guid;
using CLSID = System.Guid;
using DWORD = System.UInt32;

namespace ExcelDna.Integration.Rtd
{
    internal static class ExcelRtd
    {
        // Internal RTD Server support.
        // We write (and quickly remove again) entries in HKEY_USER\Software\Classes....
        // This is consistent with the ExcelComAddIn plan, where we need to register a Com Add-In.
        // Not trying CreateInstance / Moniker plans for now.

        // Map name to RtdServer.
        // DOCUMENT: We allow more than one name per server (ProgId and FullName)
        // but will make separate instances for the different names...?
        static Dictionary<string, Type> registeredRtdServerTypes = new Dictionary<string, Type>();
        // Map names of loaded Rtd servers to a registered ProgId - "RtdSrv.A1B2C3...."
        static Dictionary<string, string> loadedRtdServers = new Dictionary<string, string>();

        public static void RegisterRtdServerTypes(Dictionary<string, Type> rtdServerTypes)
        {
            // Just merge registrations into the overall list.
            foreach (string key in rtdServerTypes.Keys)
            {
                registeredRtdServerTypes[key] = rtdServerTypes[key];
            }
        }

        // Forwarded from XlCall
        public static object RTD(string progId, string server, params string[] topics)
        {
            // Check if this is any of our business.
            if (!string.IsNullOrEmpty(server) || !registeredRtdServerTypes.ContainsKey(progId))
            {
                // Just pass on to Excel.
                return CallRTD(progId, null, topics);
            }

            // Check if already loaded.
            if (loadedRtdServers.ContainsKey(progId))
            {
                // Call Excel using the RtdSrv.xxx ProgId
                return CallRTD(loadedRtdServers[progId], null, topics);
            }

            // Need to get the Rtd server loaded
            // We pick a new Guid as ClassId for this add-in...
            CLSID clsId = Guid.NewGuid();
            // ...and make the ProgId from this Guid - max 39 chars.
            string progIdRegistered = "RtdSrv." + clsId.ToString("N");

            Type rtdServerType = registeredRtdServerTypes[progId];
            object rtdServer = Activator.CreateInstance(rtdServerType);
            RtdServerWrapper rtdServerWrapper = new RtdServerWrapper(rtdServer, progId);

            // Mark as loaded - ServerTerminate in the wrapper will remove.
            // TODO: Consider multithread race condition...
            loadedRtdServers[progId] = progIdRegistered;
            using (SingletonClassFactoryRegistration regClassFactory = new SingletonClassFactoryRegistration(clsId, rtdServerWrapper))
            {
                using (ProgIdRegistration regProgId = new ProgIdRegistration(progIdRegistered, clsId))
                {
                    using (ClsIdRegistration regClsId = new ClsIdRegistration(clsId))
                    {
                        return CallRTD(progIdRegistered, null, topics);
                    }
                }
            }
        }

        private static object CallRTD(string progId, string server, params string[] topics)
        {
            object result;
            object[] args = new object[topics.Length + 2];
            args[0] = progId;
            args[1] = null;
            topics.CopyTo(args, 2);
            XlCall.XlReturn retval = XlCall.TryExcel(XlCall.xlfRtd, out result, args);
            if (retval == XlCall.XlReturn.XlReturnSuccess)
            {
                return result;
            }
            Debug.Print("RTD Call failed. Excel returned {0}", retval);
            return null;
        }

        public static void UnregisterRTDServer(string progId)
        {
            loadedRtdServers.Remove(progId);
        }
    }

    [ClassInterface(ClassInterfaceType.None)]
    internal class RtdServerWrapper : IRtdServer
    {
        // 'ProgId' under which ExcelDna registered the server 
        // - might be the class FullName or might be the ProgIdAttribute.
        string _progId;
        
        // If the object implements ExcelDna.Integration.Rtd.IRtdServer we call directly...
        IRtdServer _rtdServer;  
        // ... otherwise we call via delegates assigned through interface mapping.
        // Private delegate types for the IRtdServer interface ...
        delegate object delConnectData(int topicId, ref Array strings, ref bool newValues);
        delegate void delDisconnectData(int topicId);
        delegate int delHeartbeat();
        delegate Array delRefreshData(ref int topicCount);
        delegate int delServerStart(IRTDUpdateEvent CallbackObject); // Careful - might be an unexpected IRTDUpdateEvent...
        delegate void delServerTerminate();

        // ... and corresponding instances.
        delConnectData      _ConnectData;
        delDisconnectData   _DisconnectData;
        delHeartbeat        _Heartbeat;
        delRefreshData      _RefreshData;
        delServerStart      _ServerStart;
        delServerTerminate  _ServerTerminate;

        public RtdServerWrapper(object rtdServer, string progId)
        {
            _progId = progId;
            _rtdServer = rtdServer as IRtdServer;
            if (_rtdServer == null)
            {
                // The RtdServer implements another instance of IRtdServer (maybe from some office PIA).
                // We put together some delegates to call through.
                Type[] itfs = rtdServer.GetType().GetInterfaces();
                foreach (Type itf in itfs)
                {
                    if (itf.GUID == ComAPI.guidIRtdServer)
                    {
                        InterfaceMapping map = rtdServer.GetType().GetInterfaceMap(itf);
                        for (int i = 0; i < map.InterfaceMethods.Length; i++)
                        {
                            MethodInfo mi = map.InterfaceMethods[i];
                            switch (mi.Name)
                            {
                                case "ConnectData":
                                    _ConnectData = (delConnectData)Delegate.CreateDelegate(typeof(delConnectData), rtdServer, map.TargetMethods[i]);
                                    break;
                                case "DisconnectData":
                                    _DisconnectData = (delDisconnectData)Delegate.CreateDelegate(typeof(delDisconnectData), rtdServer, map.TargetMethods[i]);
                                    break;
                                case "Heartbeat":
                                    _Heartbeat = (delHeartbeat)Delegate.CreateDelegate(typeof(delHeartbeat), rtdServer, map.TargetMethods[i]);
                                    break;
                                case "RefreshData":
                                    _RefreshData = (delRefreshData)Delegate.CreateDelegate(typeof(delRefreshData), rtdServer, map.TargetMethods[i]);
                                    break;
                                case "ServerStart":
                                    // ServerStart is tricky because of the parameter type mapping.
                                    MethodInfo serverStartMethod = map.TargetMethods[i];
                                    _ServerStart = delegate(IRTDUpdateEvent updateEvent)
                                    {
                                        return (int)serverStartMethod.Invoke(rtdServer, new object[] {updateEvent});
                                    };
                                    break;
                                case "ServerTerminate":
                                    _ServerTerminate = (delServerTerminate)Delegate.CreateDelegate(typeof(delServerTerminate), rtdServer, map.TargetMethods[i]);
                                    break;
                            }
                        }
                    }
                }
            }
        }

        public object ConnectData(int topicId, ref Array strings, ref bool newValues)
        {
            if (_rtdServer != null)
            {
                return _rtdServer.ConnectData(topicId, ref strings, ref newValues);
            }
            return _ConnectData(topicId, ref strings, ref newValues);
        }

        public void DisconnectData(int topicId)
        {
            if (_rtdServer != null)
            {
                _rtdServer.DisconnectData(topicId);
                return;
            }
            _DisconnectData(topicId);
        }

        public int Heartbeat()
        {
            if (_rtdServer != null)
            {
                return _rtdServer.Heartbeat();
            }
            return _Heartbeat();
        }

        public Array RefreshData(ref int topicCount)
        {
            if (_rtdServer != null)
            {
                return _rtdServer.RefreshData(ref topicCount);
            }
            return _RefreshData(ref topicCount);
        }

        public int ServerStart(IRTDUpdateEvent CallbackObject)
        {
            if (_rtdServer != null)
            {
                return _rtdServer.ServerStart(CallbackObject);
            }
            // CallbackObject will actually be a RCW (__ComObject) so the type 'mismatch' calling Invoke never arises.
            return _ServerStart(CallbackObject);
        }

        public void ServerTerminate()
        {
            ExcelRtd.UnregisterRTDServer(_progId);
            if (_rtdServer != null)
            {
                _rtdServer.ServerTerminate();
                return;
            }
            _ServerTerminate();
        }
    }
}
