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
            throw new NotImplementedException();
        }

        int IDispatch.GetTypeInfo(uint iTInfo, uint lcid, out nint ppTInfo)
        {
            throw new NotImplementedException();
        }

        int IDispatch.GetIDsOfNames(Guid riid, string[] rgszNames, uint cNames, uint lcid, int[] rgDispId)
        {
            throw new NotImplementedException();
        }

        int IDispatch.Invoke(int dispIdMember, Guid riid, uint lcid, INVOKEKIND wFlags, in DispParams pDispParams, nint pVarResult, nint pExcepInfo, nint puArgErr)
        {
            throw new NotImplementedException();
        }
    }
}

#endif
