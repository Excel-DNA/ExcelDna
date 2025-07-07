#if COM_GENERATED

using ExcelDna.Integration.Extensibility;
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
{
    [GeneratedComInterface]
    [Guid(ExcelDna.ComInterop.ComAPI.gstrIDTExtensibility2)]
    internal partial interface IDTExtensibility2 : IDispatch
    {
        void OnConnection(IntPtr Application, ext_ConnectMode ConnectMode, IntPtr AddInInst, ref SafeArray custom);

        void OnDisconnection(ext_DisconnectMode RemoveMode, ref SafeArray custom);

        void OnAddInsUpdate(ref SafeArray custom);

        void OnStartupComplete(ref SafeArray custom);

        void OnBeginShutdown(ref SafeArray custom);
    }
}

#endif
