#if COM_GENERATED

using System;
using System.Runtime.InteropServices.Marshalling;

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
{
    internal class UnknownObject
    {
        public IntPtr P { get; }

        public UnknownObject(IntPtr unknown)
        {
            P = unknown;
        }

        public unsafe bool HasInterface(ref Guid guid)
        {
            StrategyBasedComWrappers.DefaultIUnknownStrategy.QueryInterface(P.ToPointer(), in guid, out void* ppObj);
            return ppObj != null;
        }

        public unsafe int QueryInterface(ref Guid guid, out IntPtr ppv)
        {
            int result = StrategyBasedComWrappers.DefaultIUnknownStrategy.QueryInterface(P.ToPointer(), in guid, out void* ppObj);
            if (result == 0)
                ppv = new IntPtr(ppObj);
            else
                ppv = IntPtr.Zero;

            return result;
        }
    }
}

#endif
