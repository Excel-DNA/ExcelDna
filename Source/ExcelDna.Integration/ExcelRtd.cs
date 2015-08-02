//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;
using ExcelDna.ComInterop;
using ExcelDna.ComInterop.ComRegistration;

using HRESULT = System.Int32;
using IID = System.Guid;
using CLSID = System.Guid;
using DWORD = System.UInt32;
using ExcelDna.Logging;

namespace ExcelDna.Integration.Rtd
{
    // Internal RTD registration support. Takes care of:
    // - On-demand registration of RTD server (used via XlCall.RTD...)
    // - Provides a wrapper for the RTD server to allow us to use either the internally defined IRtdServer types 
    //   or the 'real' COM interop types. 
    //   CONSIDER: This would not be needed for .NET 4.

    // NOTE: We don't use this(?) for the ExcelObservableRtdServer 
    //       - that case is directly registered with the ComServer, and in the registry via AsyncUtil.Register()

    internal static class RtdRegistration
    {
        // We write (and quickly remove again) entries in HKEY_CURRENT_USER\Software\Classes....
        // This is consistent with the ExcelComAddIn plan, where we need to register a Com Add-In.
        // Not trying CreateInstance / Moniker plans for now (any help appreciated...)

        // Map name to RtdServer.
        // DOCUMENT: We allow more than one name per server (ProgId and FullName)
        // but will make separate instances for the different names...?
        static readonly Dictionary<string, Type> registeredRtdServerTypes = new Dictionary<string, Type>();
        // Map names of loaded Rtd servers to a registered ProgId - "RtdSrv.A1B2C3...."
        static readonly Dictionary<string, string> loadedRtdServers = new Dictionary<string, string>();

        public static void RegisterRtdServerTypes(IEnumerable<Type> rtdServerTypes)
        {
            foreach (Type rtdType in rtdServerTypes)
            {
                // Decide under what name(s) to register.
                object[] attrs = rtdType.GetCustomAttributes(typeof(ProgIdAttribute), false);
                if (attrs.Length >= 1)
                {
                    ProgIdAttribute progIdAtt = (ProgIdAttribute)attrs[0];
                    registeredRtdServerTypes[progIdAtt.Value] = rtdType;
                    Logger.Initialization.Verbose("RTD Server found - Type {0}, ProgId {1}", rtdType.FullName, progIdAtt.Value);
                }
                else
                {
                    registeredRtdServerTypes[rtdType.FullName] = rtdType;
                    Logger.Initialization.Verbose("RTD Server found - Type {0} (No ProgId)", rtdType.FullName);
                }
            }
        }

        // Forwarded from XlCall
        // Loads the RTD server with temporary ProgId.
        // CAUTION: Might fail when called from array formula (the first call in every array-group fails).
        //          When it fails, the xlfRtd call returns xlReturnUncalced.
        //          In that case, this function returns null, and does not keep a reference to the created server object.
        //          The next call should then succeed (though a new server object will be created).
        public static bool TryRTD(out object result, string progId, string server, params string[] topics)
        {
            Debug.Print("### RtdRegistration.RTD " + progId);
            // Check if this is any of our business.
            Type rtdServerType;
            if (!string.IsNullOrEmpty(server) || !registeredRtdServerTypes.TryGetValue(progId, out rtdServerType))
            {
                // Just pass on to Excel.
                return TryCallRTD(out result, progId, null, topics);
            }

            // TODO: Check that ExcelRtdServer with stable ProgId case also works right here - 
            //       might need to add to loadedRtdServers somehow

            // Check if already loaded.
            string loadedProgId;
            if (loadedRtdServers.TryGetValue(progId, out loadedProgId))
            {
                if (ExcelRtd2010BugHelper.ExcelVersionHasRtdBug && rtdServerType.IsSubclassOf(typeof(ExcelRtdServer)))
                {
                    ExcelRtd2010BugHelper.RecordRtdCall(progId, topics);
                }
                // Call Excel using the synthetic RtdSrv.xxx (or actual from attribute) ProgId
                return TryCallRTD(out result, loadedProgId, null, topics);
            }

            // Not loaded already - need to get the Rtd server loaded
            // TODO: Need to reconsider registration here.....
            //       Sometimes need stable ProgIds.
            object rtdServer;
            if (ExcelRtd2010BugHelper.ExcelVersionHasRtdBug && rtdServerType.IsSubclassOf(typeof(ExcelRtdServer)))
            {
                Debug.Print("### Creating Wrapper " + progId);
                rtdServer = new ExcelRtd2010BugHelper(progId, rtdServerType);
            }
            else
            {
                using (XlCall.Suspend())
                {
                    rtdServer = Activator.CreateInstance(rtdServerType);
                }
                ExcelRtdServer excelRtdServer = rtdServer as ExcelRtdServer;
                if (excelRtdServer != null)
                {
                    // Set ProgId so that it can be 'unregistered' (removed from loadedRtdServers) when the RTD server terminates.
                    excelRtdServer.RegisteredProgId = progId;
                }
                else
                {
                    // Make a wrapper if we are not an ExcelRtdServer
                    // (ExcelRtdServer implements exception-handling and XLCall supension itself)
                    rtdServer = new RtdServerWrapper(rtdServer, progId);
                }
            }
            
            // We pick a new Guid as ClassId for this add-in...
            CLSID clsId = Guid.NewGuid();
            
            // ... (bad idea - this will cause Excel to try to load this RTD server while it is not registered.)
            // Guid typeGuid = GuidUtilit.CreateGuid(..., DnaLibrary.XllPath + ":" + rtdServerType.FullName);
            // or something based on ExcelDnaUtil.XllGuid
            // string progIdRegistered = "RtdSrv." + typeGuid.ToString("N");

            // by making a fresh progId, we are sure Excel will try to load when we are ready.
            string progIdRegistered = "RtdSrv." + clsId.ToString("N");
            Debug.Print("RTD - Using ProgId: {0} for type: {1}", progIdRegistered, rtdServerType.FullName);

            try
            {
                using (new SingletonClassFactoryRegistration(rtdServer, clsId))
                using (new ProgIdRegistration(progIdRegistered, clsId))
                using (new ClsIdRegistration(clsId, progIdRegistered))
                {
                    Debug.Print("### About to call TryCallRTD " + progId);
                    if (TryCallRTD(out result, progIdRegistered, null, topics))
                    {
                        // Mark as loaded - ServerTerminate in the wrapper will remove.
                        loadedRtdServers[progId] = progIdRegistered;
                        Debug.Print("### Added to loadedRtdServers " + progId);
                        return true;
                    }
                    return false;
                }
            }
            catch (UnauthorizedAccessException secex)
            {
                Logger.RtdServer.Error("The RTD server of type {0} required by add-in {1} could not be registered.\r\nThis may be due to restricted permissions on the user's HKCU\\Software\\Classes key.\r\nError message: {2}", rtdServerType.FullName, DnaLibrary.CurrentLibrary.Name, secex.Message);
                result = ExcelErrorUtil.ToComError(ExcelError.ExcelErrorValue);
                // Return true to have the #VALUE stick, just as it was before the array-call refactoring
                return true;
            }
            catch (Exception ex)
            {
                Logger.RtdServer.Error("The RTD server of type {0} required by add-in {1} could not be registered.\r\nThis is an unexpected error.\r\nError message: {2}", rtdServerType.FullName, DnaLibrary.CurrentLibrary.Name, ex.Message);
                Debug.Print("RtdRegistration.RTD exception: " + ex.ToString());
                result = ExcelErrorUtil.ToComError(ExcelError.ExcelErrorValue);
                // Return true to have the #VALUE stick, just as it was before the array-call refactoring
                return true;
            }
        }

        // Returns true if the Excel call succeeded.
        private static bool TryCallRTD(out object result, string progId, string server, params string[] topics)
        {
            object[] args = new object[topics.Length + 2];
            args[0] = progId;
            args[1] = null;
            topics.CopyTo(args, 2);
            XlCall.XlReturn retval = XlCall.TryExcel(XlCall.xlfRtd, out result, args);
            if (retval == XlCall.XlReturn.XlReturnSuccess)
            {
                // All is good
                return true;
            }
            if (retval == XlCall.XlReturn.XlReturnUncalced)
            {
                // An expected error - the first call in an array-group seems to always return this,
                // to be followed by one call for each element in the array (where xlfRtd succceeds).
                Debug.Print("### RTD Call failed. Excel returned {0}", retval);
                result = null;
                return false;
            }
            // Unexpected error - throw for the user to deal with
            throw new XlCallException(retval);
        }

        public static void UnregisterRTDServer(string progId)
        {
            Debug.Print("### UnregisterRTDServer " + progId);
            // Dictionary.Remove is safe to call even if the key does not exist (just returns false)
            if (progId != null)
            {
                loadedRtdServers.Remove(progId);
            }
        }
    }

    // This wrapper is not used with servers implemented from ExcelRtdServer 
    [ClassInterface(ClassInterfaceType.None)]
    internal class RtdServerWrapper : IRtdServer
    {
        // 'ProgId' under which ExcelDna registered the server 
        // - might be the class FullName or might be the ProgIdAttribute.
        readonly string _progId;
        
        // If the object implements ExcelDna.Integration.Rtd.IRtdServer we call directly...
        readonly IRtdServer _rtdServer;  
        // ... otherwise we call via delegates assigned through interface mapping.
        // Private delegate types for the IRtdServer interface ...
        delegate object delConnectData(int topicId, ref Array strings, ref bool newValues);
        delegate void   delDisconnectData(int topicId);
        delegate int    delHeartbeat();
        delegate Array  delRefreshData(ref int topicCount);
        delegate int    delServerStart(IRTDUpdateEvent CallbackObject); // Careful - might be an unexpected IRTDUpdateEvent...
        delegate void   delServerTerminate();

        // ... and corresponding instances.
        readonly delConnectData     _ConnectData;
        readonly delDisconnectData  _DisconnectData;
        readonly delHeartbeat       _Heartbeat;
        readonly delRefreshData     _RefreshData;
        readonly delServerStart     _ServerStart;
        readonly delServerTerminate _ServerTerminate;

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
                using (XlCall.Suspend())
                {
                    if (_rtdServer != null)
                    {
                        return _rtdServer.ConnectData(topicId, ref strings, ref newValues);
                    }
                    return _ConnectData(topicId, ref strings, ref newValues);
                }
            }
            catch (Exception e)
            {
                Logger.RtdServer.Error("Error in RTD server {0} ConnectData: {1}", _progId, e.ToString());
                return null;
            }
        }

        public void DisconnectData(int topicId)
        {
            try
            {
                using (XlCall.Suspend())
                {
                    if (_rtdServer != null)
                    {
                        _rtdServer.DisconnectData(topicId);
                        return;
                    }
                    _DisconnectData(topicId);
                }
            }
            catch (Exception e)
            {
                Logger.RtdServer.Error("Error in RTD server {0} DisconnectData: {1}", _progId, e.ToString());
            }
        }

        public int Heartbeat()
        {
            try
            {
                using (XlCall.Suspend())
                {
                    if (_rtdServer != null)
                    {
                        return _rtdServer.Heartbeat();
                    }
                    return _Heartbeat();
                }
            }
            catch (Exception e)
            {
                Logger.RtdServer.Error("Error in RTD server {0} Heartbeat: {1}", _progId, e.ToString());
                return 0;
            }
        }

        public Array RefreshData(ref int topicCount)
        {
            try
            {
                using (XlCall.Suspend())
                {
                    if (_rtdServer != null)
                    {
                        return _rtdServer.RefreshData(ref topicCount);
                    }
                    return _RefreshData(ref topicCount);
                }
            }
            catch (Exception e)
            {
                Logger.RtdServer.Error("Error in RTD server {0} RefreshData: {1}", _progId, e.ToString());
                return null;
            }
        }

        public int ServerStart(IRTDUpdateEvent CallbackObject)
        {
            try
            {
                using (XlCall.Suspend())
                {
                    if (_rtdServer != null)
                    {
                        return _rtdServer.ServerStart(CallbackObject);
                    }
                    // CallbackObject will actually be a RCW (__ComObject) so the type 'mismatch' calling Invoke never arises.
                    return _ServerStart(CallbackObject);
                }
            }
            catch (Exception e)
            {
                Logger.RtdServer.Error("Error in RTD server {0} ServerStart: {1}", _progId, e.ToString());
                return 0;
            }
        }

        public void ServerTerminate()
        {
            try
            {
                using (XlCall.Suspend())
                {
                    RtdRegistration.UnregisterRTDServer(_progId);
                    if (_rtdServer != null)
                    {
                        _rtdServer.ServerTerminate();
                        return;
                    }
                    _ServerTerminate();
                }
            }
            catch (Exception e)
            {
                Logger.RtdServer.Error("Error in RTD server {0} ServerTerminate: {1}", _progId, e.ToString());
            }
        }
    }

}
