#if !USE_WINDOWS_FORMS

using System;
using System.Runtime.InteropServices;

namespace ExcelDna.Integration.Win32
{
    internal class InstanceHandle : SafeHandle
    {
        public InstanceHandle(IntPtr handle) : base(0, false)
        {
            if (handle == 0)
                IsInvalid = true;
            else
                SetHandle(handle);
        }

        public override bool IsInvalid { get; }

        protected override bool ReleaseHandle()
        {
            throw new NotImplementedException();
        }
    }
}

#endif
