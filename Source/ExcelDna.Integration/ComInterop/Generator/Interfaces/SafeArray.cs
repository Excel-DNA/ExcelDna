#if COM_GENERATED

using System.Runtime.CompilerServices;

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
{
    internal partial struct SAFEARRAYBOUND
    {
        public uint cElements;
        public int lLbound;
    }

    [InlineArray(InlineArraySAFEARRAYBOUND_1.Length)]
    internal partial struct InlineArraySAFEARRAYBOUND_1
    {
        public const int Length = 1;

        public SAFEARRAYBOUND Data;
    }

    internal partial struct SafeArray
    {
        public ushort cDims;
        public ushort fFeatures;
        public uint cbElements;
        public uint cLocks;
        public nint pvData;
        public InlineArraySAFEARRAYBOUND_1 rgsabound; // variable-length array placeholder
    }
}

#endif
