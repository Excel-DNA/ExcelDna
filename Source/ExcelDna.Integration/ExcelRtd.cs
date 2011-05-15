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
        // We write (and quickly remove again) entries in HKEY_CURRENT_USER\Software\Classes....
        // This is consistent with the ExcelComAddIn plan, where we need to register a Com Add-In.
        // Not trying CreateInstance / Moniker plans for now.

        // Map name to RtdServer.
        // DOCUMENT: We allow more than one name per server (ProgId and FullName)
        // but will make separate instances for the different names...?
        static Dictionary<string, Type> registeredRtdServerTypes = new Dictionary<string, Type>();
        // Map names of loaded Rtd servers to a registered ProgId - "RtdSrv.A1B2C3...."
        static Dictionary<string, string> loadedRtdServers = new Dictionary<string, string>();

        public static void RegisterRtdServerTypes(List<Type> rtdServerTypes)
        {
            foreach (Type rtdType in rtdServerTypes)
            {
                // Decide under what name(s) to register.
                object[] attrs = rtdType.GetCustomAttributes(typeof(ProgIdAttribute), false);
                if (attrs.Length >= 1)
                {
                    ProgIdAttribute progIdAtt = (ProgIdAttribute)attrs[0];
                    registeredRtdServerTypes[progIdAtt.Value] = rtdType;
                }
                registeredRtdServerTypes[rtdType.FullName] = rtdType;
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
                // Call Excel using the synthetic RtdSrv.xxx (or actual from attribute) ProgId
                return CallRTD(loadedRtdServers[progId], null, topics);
            }

            // Not loaded already - need to get the Rtd server loaded

            // TODO: Need to reconsider registration here.....
            // Sometimes need stable ProgIds.

            Type rtdServerType = registeredRtdServerTypes[progId];
            object rtdServer = Activator.CreateInstance(rtdServerType);
            RtdServerWrapper rtdServerWrapper = new RtdServerWrapper(rtdServer, progId);

            // We pick a new Guid as ClassId for this add-in...
            CLSID clsId = Guid.NewGuid();
            
            // ... but take a stable ProgId from the XllPath and type's FullName - max 39 chars.
            // ... (bad idea - this will cause Excel to try to load this RTD server while it is not registered.)
            // Guid typeGuid = GetGuidFromString(DnaLibrary.XllPath + ":" + rtdServerType.FullName);
            // string progIdRegistered = "RtdSrv." + typeGuid.ToString("N");

            // by making a fresh progId, we are sure Excel will try to load when we are ready.
            string progIdRegistered = "RtdSrv." + clsId.ToString("N");
            Debug.Print("RTD - Using ProgId: {0} for type: {1}", progIdRegistered, rtdServerType.FullName);

            // Mark as loaded - ServerTerminate in the wrapper will remove.
            // TODO: Consider multithread race condition...
            loadedRtdServers[progId] = progIdRegistered;
            using (SingletonClassFactoryRegistration regClassFactory = new SingletonClassFactoryRegistration(rtdServerWrapper, clsId))
            using (ProgIdRegistration regProgId = new ProgIdRegistration(progIdRegistered, clsId))
            using (ClsIdRegistration regClsId = new ClsIdRegistration(clsId, progIdRegistered))
            {
                return CallRTD(progIdRegistered, null, topics);
            }
        }

        //private static Guid GetGuidFromString(string input)
        //{
        //    // CONSIDER: Not sure if this is really the right way - there is a real GUID category for hash -> Guid generation

        //    //string path = @"D:\Files\CSharp\Excel\ExcelDna\SomeExampleXll.xll";
        //    //byte[] tmp = System.Text.Encoding.Unicode.GetBytes(path); 
        //    byte[] tmp = System.Text.Encoding.UTF8.GetBytes(input);
        //    byte[] hash = System.Security.Cryptography.MD5.Create().ComputeHash(tmp);
        //    Guid g = new Guid(hash);
        //    return g;
        //}

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
            if (loadedRtdServers.ContainsKey(progId))
            {
                loadedRtdServers.Remove(progId);
            }
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
            // CAREFUL: ProgId passed in might be token used from regular ClassFactory.
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
            try
            {
                if (_rtdServer != null)
                {
                    return _rtdServer.ConnectData(topicId, ref strings, ref newValues);
                }
                return _ConnectData(topicId, ref strings, ref newValues);
            }
            catch (Exception e)
            {
                Logging.LogDisplay.WriteLine("Error in RTD server {0} ConnectData: {1}", _progId, e.ToString());
                return null;
            }
        }

        public void DisconnectData(int topicId)
        {
            try
            {
                if (_rtdServer != null)
                {
                    _rtdServer.DisconnectData(topicId);
                    return;
                }
                _DisconnectData(topicId);
            }
            catch (Exception e)
            {
                Logging.LogDisplay.WriteLine("Error in RTD server {0} DisconnectData: {1}", _progId, e.ToString());
            }
        }

        public int Heartbeat()
        {
            try
            {
                if (_rtdServer != null)
                {
                    return _rtdServer.Heartbeat();
                }
                return _Heartbeat();
            }
            catch (Exception e)
            {
                Logging.LogDisplay.WriteLine("Error in RTD server {0} Heartbeat: {1}", _progId, e.ToString());
                return 0;
            }
        }

        public Array RefreshData(ref int topicCount)
        {
            try
            {
                if (_rtdServer != null)
                {
                    return _rtdServer.RefreshData(ref topicCount);
                }
                return _RefreshData(ref topicCount);
            }
            catch (Exception e)
            {
                Logging.LogDisplay.WriteLine("Error in RTD server {0} RefreshData: {1}", _progId, e.ToString());
                return null;
            }
        }

        public int ServerStart(IRTDUpdateEvent CallbackObject)
        {
            try
            {
                if (_rtdServer != null)
                {
                    return _rtdServer.ServerStart(CallbackObject);
                }
                // CallbackObject will actually be a RCW (__ComObject) so the type 'mismatch' calling Invoke never arises.
                return _ServerStart(CallbackObject);
            }
            catch (Exception e)
            {
                Logging.LogDisplay.WriteLine("Error in RTD server {0} ServerStart: {1}", _progId, e.ToString());
                return 0;
            }
        }

        public void ServerTerminate()
        {
            try
            {
                ExcelRtd.UnregisterRTDServer(_progId);
                if (_rtdServer != null)
                {
                    _rtdServer.ServerTerminate();
                    return;
                }
                _ServerTerminate();
            }
            catch (Exception e)
            {
                Logging.LogDisplay.WriteLine("Error in RTD server {0} ServerTerminate: {1}", _progId, e.ToString());
            }
        }
    }
}
