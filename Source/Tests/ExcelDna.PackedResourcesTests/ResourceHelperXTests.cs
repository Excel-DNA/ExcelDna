using ExcelDna.PackedResources;
using NUnit.Framework;
using System.IO;

namespace ExcelDna.PackedResourcesTests
{
    public class ResourceHelperXTests
    {
        [Test]
        public void AddResource()
        {
            string outPath = TestdataHelper.FilePath("ResourceHelperXTests-AddIn64-packedX-out.dll");
            File.Copy(TestdataHelper.FilePath("AddIn64.dll"), outPath, true);
            ResourceHelperX.AddResource(outPath, File.ReadAllBytes(TestdataHelper.FilePath("test_pack-AddIn.dna.bin")), "__MAIN__", "DNA");
            ResourceHelperX.AddResource(outPath, File.ReadAllBytes(TestdataHelper.FilePath("test_pack.dll.bin")), "TEST_PACK", "ASSEMBLY_LZMA");

            Assert.That(File.ReadAllBytes(outPath), Is.EqualTo(File.ReadAllBytes(TestdataHelper.FilePath("AddIn64-packedX.dll"))));
            File.Delete(outPath);
        }
    }
}
