#if COM_GENERATED

using ExcelDna.Integration.ComInterop.Generator.Interfaces;
using System;
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
            dispatcher.Invoke(dispIdMember, pDispParams);

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

        //

        private void OnServerStart(DispParams pDispParams)
        {
            IRTDUpdateEvent callbackObject = (pDispParams.rgvarg[0].Value as DispatchObject).ComObject as IRTDUpdateEvent;
        }

        private void OnServerTerminate(DispParams pDispParams)
        {
        }
    }
}

#endif
