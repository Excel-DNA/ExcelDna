using System;

namespace ExcelDna.PackedResources
{
    internal interface IResourceResolver
    {
        void Begin(string fileName);
        void End(bool discard);
        bool Update(string lpType, IntPtr intResource, ushort wLanguage, IntPtr lpData, uint cbData);
        bool Update(string lpType, string lpName, ushort wLanguage, byte[] data);
        bool Update(uint lpType, uint lpName, ushort wLanguage, IntPtr lpData, uint cbData);
    }
}
