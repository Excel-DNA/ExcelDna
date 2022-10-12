using System.IO;

namespace ExcelDna.PackedResourcesTests
{
    internal class TestdataHelper
    {
        public static string FilePath(string relativePath)
        {
            return Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "testdata", relativePath);
        }
    }
}
