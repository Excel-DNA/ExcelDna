#if COM_GENERATED

using ExcelDna.Integration.ComInterop.Generator.Interfaces;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices.Marshalling;

namespace ExcelDna.Integration.ComInterop.Generator
{
    [GeneratedComClass]
    internal partial class ExcelObserverRtdServer : Rtd.ExcelObserverRtdServer, IRtdServer
    {
        private Dispatcher dispatcher;

        public ExcelObserverRtdServer()
        {
            dispatcher = new Dispatcher(new Dispatcher.Method[] {
                new Dispatcher.Method("ServerStart", OnServerStart),
                new Dispatcher.Method("ConnectData", OnConnectData),
                new Dispatcher.Method("RefreshData", OnRefreshData),
                new Dispatcher.Method("DisconnectData", OnDisconnectData),
                new Dispatcher.Method("Heartbeat", OnHeartbeat),
                new Dispatcher.Method("ServerTerminate", OnServerTerminate),
            });
        }

        // IDispatch:
        public int GetTypeInfoCount(out uint pctinfo)
        {
            throw new NotImplementedException();
        }

        public int GetTypeInfo(uint iTInfo, uint lcid, out nint ppTInfo)
        {
            throw new NotImplementedException();
        }

        public int GetIDsOfNames(Guid riid, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 2)] string[] rgszNames, uint cNames, uint lcid, [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2), Out] int[] rgDispId)
        {
            dispatcher.GetIDsOfNames(rgszNames, rgDispId);

            return 0;
        }

        public int Invoke(int dispIdMember, Guid riid, uint lcid, INVOKEKIND wFlags, [MarshalUsing(typeof(DispParamsMarshaller))] in DispParams pDispParams, nint pVarResult, nint pExcepInfo, nint puArgErr)
        {
            dispatcher.Invoke(dispIdMember, pDispParams, pVarResult);

            return 0;
        }

        // IRtdServer:
        public int ServerStart(nint CallbackObject)
        {
            throw new NotImplementedException();
        }

        public nint ConnectData(int topicId, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] strings, ref int newValues)
        {
            throw new NotImplementedException();
        }

        public SafeArray RefreshData(ref int topicCount)
        {
            throw new NotImplementedException();
        }

        public void DisconnectData(int topicID)
        {
            throw new NotImplementedException();
        }

        public int Heartbeat()
        {
            throw new NotImplementedException();
        }

        public void ServerTerminate()
        {
            throw new NotImplementedException();
        }

        // IRtdServer adapter:

        private void OnServerStart(DispParams pDispParams, nint pVarResult)
        {
            IRTDUpdateEvent callbackObject = (pDispParams.rgvarg[0].Value as DispatchObject).ComObject as IRTDUpdateEvent;
            Dispatcher.SetResult(pVarResult, (this as Rtd.IRtdServer).ServerStart(new RTDUpdateEvent(callbackObject)));
        }

        private void OnConnectData(DispParams pDispParams, nint pVarResult)
        {
            Variant[] vstrings = pDispParams.rgvarg[1].Value as Variant[];
            Array strings = vstrings.Select(i => i.Value as string).ToArray();
            bool newValues = (bool)pDispParams.rgvarg[2].Value;
            Dispatcher.SetResult(pVarResult, (this as Rtd.IRtdServer).ConnectData(
                (int)pDispParams.rgvarg[0].Value,
                ref strings,
                ref newValues));
            DispParamsMarshaller.UpdateRefBoolArg(pDispParams, newValues, 2);
        }

        private void OnRefreshData(DispParams pDispParams, nint pVarResult)
        {
            int topicCount = (int)pDispParams.rgvarg[0].Value;
            Dispatcher.SetResult(pVarResult, (this as Rtd.IRtdServer).RefreshData(ref topicCount));
            DispParamsMarshaller.UpdateRefIntArg(pDispParams, topicCount, 0);
        }

        private void OnDisconnectData(DispParams pDispParams, nint pVarResult)
        {
            (this as Rtd.IRtdServer).DisconnectData((int)pDispParams.rgvarg[0].Value);
        }

        private void OnHeartbeat(DispParams pDispParams, nint pVarResult)
        {
            Dispatcher.SetResult(pVarResult, (this as Rtd.IRtdServer).Heartbeat());
        }

        private void OnServerTerminate(DispParams pDispParams, nint pVarResult)
        {
            (this as Rtd.IRtdServer).ServerTerminate();
        }
    }
}

#endif
