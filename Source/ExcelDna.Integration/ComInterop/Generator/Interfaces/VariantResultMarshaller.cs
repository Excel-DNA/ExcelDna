#if COM_GENERATED

using System;
using System.Runtime.InteropServices;

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
{
    internal class VariantResultMarshaller : IDisposable
    {
        private bool disposedValue;

        public nint Ptr { get; }

        public VariantResultMarshaller()
        {
            Ptr = Marshal.AllocHGlobal(Marshal.SizeOf<VariantNative>());
            Marshal.StructureToPtr(default(VariantNative), Ptr, false);
        }

        public Variant GetResult()
        {
            VariantNative variantNative = Marshal.PtrToStructure<VariantNative>(Ptr);
            return VariantMarshaller.ConvertToManaged(variantNative);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                VariantMarshaller.Free(Marshal.PtrToStructure<VariantNative>(Ptr));
                Marshal.FreeHGlobal(Ptr);
                disposedValue = true;
            }
        }

        ~VariantResultMarshaller()
        {
            Dispose(disposing: false);
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}

#endif
