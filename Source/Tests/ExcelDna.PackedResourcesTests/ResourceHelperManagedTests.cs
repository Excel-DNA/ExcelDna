using ExcelDna.PackedResources;
using NUnit.Framework;
using System.IO;

namespace ExcelDna.PackedResourcesTests
{
    public class ResourceHelperManagedTests
    {
        [Test]
        public void AddResource()
        {
            string outPath = TestdataHelper.FilePath("ResourceHelperManagedTests-AddResource-out.dll");
            File.Copy(TestdataHelper.FilePath("AddIn64.dll"), outPath, true);
            ResourceHelperManaged.AddResource(outPath, File.ReadAllBytes(TestdataHelper.FilePath("test_pack-AddIn.dna.bin")), "__MAIN__", "DNA");
            ResourceHelperManaged.AddResource(outPath, File.ReadAllBytes(TestdataHelper.FilePath("test_pack.dll.bin")), "TEST_PACK", "ASSEMBLY_LZMA");

            Assert.That(File.ReadAllBytes(outPath), Is.EqualTo(File.ReadAllBytes(TestdataHelper.FilePath("AddIn64-packedX.dll"))));
            File.Delete(outPath);
        }

        [Test]
        public void AddResourceOutOfOrder()
        {
            string outPath = TestdataHelper.FilePath("ResourceHelperManagedTests-AddResourceOutOfOrder-out.dll");
            File.Copy(TestdataHelper.FilePath("AddIn64.dll"), outPath, true);
            ResourceHelperManaged.AddResource(outPath, File.ReadAllBytes(TestdataHelper.FilePath("atest_pack-AddIn.dna.bin")), "__MAIN__", "DNA");
            ResourceHelperManaged.AddResource(outPath, File.ReadAllBytes(TestdataHelper.FilePath("atest_pack.dll.bin")), "ATEST_PACK", "ASSEMBLY_LZMA");

            Assert.That(File.ReadAllBytes(outPath), Is.EqualTo(File.ReadAllBytes(TestdataHelper.FilePath("AddIn64-packedXa.dll"))));
            File.Delete(outPath);
        }
    }
}
