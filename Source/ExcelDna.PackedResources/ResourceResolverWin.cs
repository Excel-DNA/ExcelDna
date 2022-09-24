using System.Runtime.InteropServices;
using System;
using System.ComponentModel;
using System.Collections.Generic;

namespace ExcelDna.PackedResources
{
    internal class ResourceResolverWin : IResourceResolver
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr BeginUpdateResource(string pFileName, bool bDeleteExistingResources);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool EndUpdateResource(IntPtr hUpdate, bool fDiscard);

        //, EntryPoint="UpdateResourceA", CharSet=CharSet.Ansi,
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool UpdateResource(
            IntPtr hUpdate,
            string lpType,
            string lpName,
            ushort wLanguage,
            IntPtr lpData,
            uint cbData);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool UpdateResource(
            IntPtr hUpdate,
            string lpType,
            IntPtr intResource,
            ushort wLanguage,
            IntPtr lpData,
            uint cbData);

        // This overload provides the resource type and name conversions that would be done by MAKEINTRESOURCE
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool UpdateResource(
            IntPtr hUpdate,
            uint lpType,
            uint lpName,
            ushort wLanguage,
            IntPtr lpData,
            uint cbData);

        private IntPtr _hUpdate;
        private List<object> updateData = new List<object>();

        public void Begin(string fileName)
        {
            _hUpdate = BeginUpdateResource(fileName, false);
            if (_hUpdate == IntPtr.Zero)
            {
                throw new Win32Exception();
            }
        }

        public bool Update(string lpType, IntPtr intResource, ushort wLanguage, IntPtr lpData, uint cbData)
        {
            return UpdateResource(_hUpdate, lpType, intResource, wLanguage, lpData, cbData);
        }

        public bool Update(string lpType, string lpName, ushort wLanguage, byte[] data)
        {
            IntPtr lpData;
            uint cbData;
            if (data == null)
            {
                lpData = IntPtr.Zero;
                cbData = 0;
            }
            else
            {
                GCHandle pinHandle = GCHandle.Alloc(data, GCHandleType.Pinned);
                updateData.Add(pinHandle);
                lpData = pinHandle.AddrOfPinnedObject();
                cbData = (uint)data.Length;
            }

            return UpdateResource(_hUpdate, lpType, lpName, wLanguage, lpData, cbData);
        }

        public bool Update(uint lpType, uint lpName, ushort wLanguage, IntPtr lpData, uint cbData)
        {
            return UpdateResource(_hUpdate, lpType, lpName, wLanguage, lpData, cbData);
        }

        public void End(bool discard)
        {
            bool result = EndUpdateResource(_hUpdate, discard);
            if (!result)
            {
                throw new Win32Exception();
            }
        }
    }
}
