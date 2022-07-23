using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SDKExcelDnaPackXllNameTests : IntegrationTestBase
    {
        [Test]
        public void Packed()
        {
            const string projectBasePath = @"SDKExcelDnaPackXllName\";
            const string projectOutDir = projectBasePath + @"bin\Release\";
            const string publishDir = projectOutDir + @"publish\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelDnaPackXllName.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(publishDir, "*.xll", "MyPacked32Name.xll", "My64PackedName.xll");
        }
    }
}
