#if ASMRESOLVER

using System;

namespace ExcelDna.PackedResources
{
    internal class ResourceResolverManaged : IResourceResolver
    {
        private string fileName;

        public void Begin(string fileName)
        {
            this.fileName = fileName;
        }

        public bool Update(string lpType, IntPtr intResource, ushort wLanguage, IntPtr lpData, uint cbData)
        {
            throw new NotImplementedException();
        }

        public bool Update(string lpType, string lpName, ushort wLanguage, byte[] data)
        {
            if (data == null)
                throw new NotImplementedException();

            ResourceHelperManaged.AddResource(fileName, data, lpName, lpType);
            return true;
        }

        public bool Update(uint lpType, uint lpName, ushort wLanguage, IntPtr lpData, uint cbData)
        {
            throw new NotImplementedException();
        }

        public void End(bool discard)
        {
            if (discard)
                throw new NotImplementedException();
        }
    }
}
#endif
