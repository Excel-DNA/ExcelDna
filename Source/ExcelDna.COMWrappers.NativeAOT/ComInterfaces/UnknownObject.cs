using System.Runtime.InteropServices.Marshalling;

namespace ExcelDna.COMWrappers.NativeAOT.ComInterfaces
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
    }
}
