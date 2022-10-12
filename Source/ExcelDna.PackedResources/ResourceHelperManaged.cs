#if ASMRESOLVER

using System;
using System.Collections.Generic;
using System.Linq;
using AsmResolver.PE;
using AsmResolver.PE.Win32Resources;
using AsmResolver.PE.File;
using AsmResolver;
using AsmResolver.PE.Win32Resources.Builder;

namespace ExcelDna.PackedResources
{
    internal class ResourceHelperManaged
    {
        public static void AddResource(string dll, byte[] resource, string name, string dir)
        {
            var peFile = PEFile.FromFile(dll);
            var peImage = PEImage.FromFile(peFile);

            ResourceDirectory existingDirDirectory = peImage.Resources.Entries.FirstOrDefault(i => i.Name == dir) as ResourceDirectory;
            ResourceDirectory dirDirectory = existingDirDirectory ?? new ResourceDirectory(dir);

            uint uniqueIndex = 0;
            while (dirDirectory.TryGetEntry(uniqueIndex, out _))
            {
                ++uniqueIndex;
            }
            ResourceDirectory nameDirectory = new ResourceDirectory(name);
            nameDirectory.Id = uniqueIndex;

            var data = new ResourceData(0, contents: new DataSegment(resource));
            data.CodePage = 1252;
            nameDirectory.AddOrReplaceEntry(data);
            InsertInOrder(dirDirectory.Entries, nameDirectory);

            if (existingDirDirectory == null)
                InsertInOrder(peImage.Resources.Entries, dirDirectory);

            var resourceDirectoryBuffer = new ResourceDirectoryBuffer();
            resourceDirectoryBuffer.AddDirectory(peImage.Resources);

            PESection rsrc = peFile.Sections.First(x => x.Name == ".rsrc");
            rsrc.Contents = resourceDirectoryBuffer;
            peFile.Write(dll);
        }

        private static void InsertInOrder(IList<IResourceEntry> entries, IResourceEntry entry)
        {
            int i = entries.TakeWhile(i => i.Name != null && string.Compare(entry.Name, i.Name) > 0).Count();
            entries.Insert(i, entry);
        }
    }
}
#endif
