using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SDKExcelDnaPublishPathTests : IntegrationTestBase
    {
        [Test]
        public void Packed()
        {
            const string projectBasePath = @"SDKExcelDnaPublishPath\";
            const string projectOutDir = projectBasePath + @"bin\Release\";
            const string publishDir = projectOutDir + @"MyPublish\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelDnaPublishPath.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(publishDir, "*.xll", "SDKExcelDnaPublishPath-AddIn64-packed.xll", "SDKExcelDnaPublishPath-AddIn-packed.xll");
        }
    }
}
