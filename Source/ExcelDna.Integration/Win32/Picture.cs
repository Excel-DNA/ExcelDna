#if !USE_WINDOWS_FORMS

using System;
using System.Runtime.InteropServices;

namespace ExcelDna.Integration.Win32
{
    internal class Picture
    {
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void OleLoadPicture(
            nint pStream,
            int lSize,
            bool fRunMode,
            [In] ref Guid riid,
            out nint ppvObj);

        [DllImport("ole32.dll", PreserveSig = false)]
        private static extern void CreateStreamOnHGlobal(
            IntPtr hGlobal,
            bool fDeleteOnRelease,
            out nint ppstm);

        public static nint LoadAsIPictureDisp(byte[] imageData)
        {
            IntPtr hGlobal = Marshal.AllocHGlobal(imageData.Length);
            Marshal.Copy(imageData, 0, hGlobal, imageData.Length);

            nint streamPtr = IntPtr.Zero;
            CreateStreamOnHGlobal(hGlobal, true, out streamPtr);

            Guid IPictureDispGuid = new Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB");
            nint picturePtr = IntPtr.Zero;
            OleLoadPicture(streamPtr, imageData.Length, false, ref IPictureDispGuid, out picturePtr);

            return picturePtr;
        }
    }
}

#endif
