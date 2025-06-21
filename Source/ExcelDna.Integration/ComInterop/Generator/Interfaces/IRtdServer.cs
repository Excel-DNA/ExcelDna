#if COM_GENERATED

using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
{
    [GeneratedComInterface]
    [Guid(ExcelDna.ComInterop.ComAPI.gstrIRtdServer)]
    internal partial interface IRtdServer : IDispatch
    {
        int ServerStart(nint CallbackObject);
        nint ConnectData(int topicId, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] strings, ref int newValues);
        SafeArray RefreshData(ref int topicCount);
        void DisconnectData(int topicID);
        int Heartbeat();
        void ServerTerminate();
    }
}

#endif
