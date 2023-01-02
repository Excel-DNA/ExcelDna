using ExcelDna.PackedResources;
using ExcelDna.PackedResources.Logging;
using FakeItEasy;
using NUnit.Framework;
using System.IO;

namespace ExcelDna.PackedResourcesTests
{
    public class ExcelDnaPackTests
    {
        [Test]
        [TestCase(false)]
        [TestCase(true)]
        public void Pack(bool useManagedResourceResolver)
        {
            string dnaFile = TestdataHelper.FilePath("test_pack-AddIn.dna");
            string xllFile = TestdataHelper.FilePath("test_pack-AddIn.xll");
            string outPath = TestdataHelper.FilePath("ExcelDnaPackTests-AddIn64-packed-out.dll");
            File.Copy(TestdataHelper.FilePath("AddIn64-compressed.dll"), xllFile, true);
            ExcelDnaPack.Pack(dnaFile, outPath, true, false, true, null, null, false, false, null, useManagedResourceResolver, null, A.Dummy<IBuildLogger>());

            Assert.That(File.ReadAllBytes(outPath), Is.EqualTo(File.ReadAllBytes(TestdataHelper.FilePath(useManagedResourceResolver ? "AddIn64-packedX-compressed.dll" : "AddIn64-packed-compressed.dll"))));
            File.Delete(outPath);
            File.Delete(xllFile);
        }
    }
}
