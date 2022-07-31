using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SDKExcelDnaPackXllSuffixNoneTests : IntegrationTestBase
    {
        [Test]
        public void Packed()
        {
            const string projectBasePath = @"SDKExcelDnaPackXllSuffixNone\";
            const string projectOutDir = projectBasePath + @"bin\Release\";
            const string publishDir = projectOutDir + @"publish\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelDnaPackXllSuffixNone.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(publishDir, "*.xll", "SDKExcelDnaPackXllSuffixNone-AddIn64.xll", "SDKExcelDnaPackXllSuffixNone-AddIn.xll");
        }
    }
}
