#if COM_GENERATED

using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
{
    [GeneratedComInterface]
    [Guid(ExcelDna.ComInterop.ComAPI.gstrIRTDUpdateEvent)]
    internal partial interface IRTDUpdateEvent : IDispatch
    {
        void UpdateNotify();

        int get_HeartbeatInterval();

        void set_HeartbeatInterval(int value);

        void Disconnect();
    }
}

#endif
