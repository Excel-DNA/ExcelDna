using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SDKExcelDnaPublishPathNoneTests : IntegrationTestBase
    {
        [Test]
        public void Packed()
        {
            const string projectBasePath = @"SDKExcelDnaPublishPathNone\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelDnaPublishPathNone.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.xll", "SDKExcelDnaPublishPathNone-AddIn64-packed.xll", "SDKExcelDnaPublishPathNone-AddIn-packed.xll", "SDKExcelDnaPublishPathNone-AddIn64.xll", "SDKExcelDnaPublishPathNone-AddIn.xll");
        }

        [Test]
        public void Unpacked()
        {
            const string projectBasePath = @"SDKExcelDnaPublishPathNone\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelDnaPublishPathNone.csproj /t:Restore,Build /p:Configuration=Release /p:ExcelDnaUnpack=true /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dll", "ExcelDna.Integration.dll", "ExcelDna.Loader.dll", "ExcelDna.ManagedHost.dll", "SDKExcelDnaPublishPathNone.dll");
        }
    }
}
