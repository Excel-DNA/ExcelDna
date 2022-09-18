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
            string outPath = TestdataFilePath("AddIn64-packed-out.dll");
            File.Copy(TestdataFilePath("AddIn64.dll"), outPath, true);
            ResourceHelperX.AddResource(outPath, File.ReadAllBytes(TestdataFilePath("test_pack-AddIn.dna.bin")), "__MAIN__", "DNA");
            ResourceHelperX.AddResource(outPath, File.ReadAllBytes(TestdataFilePath("test_pack.dll.bin")), "TEST_PACK", "ASSEMBLY_LZMA");

            Assert.That(File.ReadAllBytes(outPath), Is.EqualTo(File.ReadAllBytes(TestdataFilePath("AddIn64-packedX.dll"))));
            File.Delete(outPath);
        }

        private static string TestdataFilePath(string relativePath)
        {
            return Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "testdata", relativePath);
        }
    }
}
