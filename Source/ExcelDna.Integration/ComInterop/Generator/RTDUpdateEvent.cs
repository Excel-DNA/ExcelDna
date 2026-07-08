#if COM_GENERATED

using ExcelDna.Integration.ComInterop.Generator.Interfaces;
using System;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices.Marshalling;

namespace ExcelDna.Integration.ComInterop.Generator
{
    [GeneratedComClass]
    internal partial class RTDUpdateEvent : Rtd.IRTDUpdateEvent, IRTDUpdateEvent
    {
        private const int S_OK = 0;
        private const int E_NOTIMPL = unchecked((int)0x80004001);

        private IRTDUpdateEvent impl;

        public RTDUpdateEvent(IRTDUpdateEvent impl)
        {
            this.impl = impl;
        }

        // Rtd.IRTDUpdateEvent:

        void Rtd.IRTDUpdateEvent.UpdateNotify()
        {
            impl.UpdateNotify();
        }

        int Rtd.IRTDUpdateEvent.HeartbeatInterval
        {
            get => throw new NotImplementedException();
            set => throw new NotImplementedException();
        }

        void Rtd.IRTDUpdateEvent.Disconnect()
        {
            throw new NotImplementedException();
        }

        // IRTDUpdateEvent:

        void IRTDUpdateEvent.UpdateNotify()
        {
            throw new NotImplementedException();
        }

        int IRTDUpdateEvent.get_HeartbeatInterval()
        {
            throw new NotImplementedException();
        }

        void IRTDUpdateEvent.set_HeartbeatInterval(int value)
        {
            throw new NotImplementedException();
        }

        void IRTDUpdateEvent.Disconnect()
        {
            throw new NotImplementedException();
        }

        // IDispatch:

        int IDispatch.GetTypeInfoCount(out uint pctinfo)
        {
            pctinfo = 0;
            return S_OK;
        }

        int IDispatch.GetTypeInfo(uint iTInfo, uint lcid, out nint ppTInfo)
        {
            ppTInfo = 0;
            return E_NOTIMPL;
        }

        int IDispatch.GetIDsOfNames(in Guid riid, string[] rgszNames, uint cNames, uint lcid, int[] rgDispId)
        {
            return E_NOTIMPL;
        }

        int IDispatch.Invoke(int dispIdMember, in Guid riid, uint lcid, ushort wFlags, in DispParams pDispParams, nint pVarResult, nint pExcepInfo, nint puArgErr)
        {
            return E_NOTIMPL;
        }
    }
}

#endif
