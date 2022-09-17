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
            ResourceHelperX.AddResource(null, null, null, null);
        }

        private static string TestdataFilePath(string relativePath)
        {
            return Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "testdata", relativePath);
        }
    }
}
