using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SDKExcelDnaPublishPathNoneTests : IntegrationTestBase
    {
        const string projectBasePath = @"SDKExcelDnaPublishPathNone\";
        const string projectOutDir = projectBasePath + @"bin\Release\";

        [Test]
        public void Packed()
        {
            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelDnaPublishPathNone.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.xll", "SDKExcelDnaPublishPathNone-AddIn64-packed.xll", "SDKExcelDnaPublishPathNone-AddIn-packed.xll", "SDKExcelDnaPublishPathNone-AddIn64.xll", "SDKExcelDnaPublishPathNone-AddIn.xll");
        }

        [Test]
        public void Unpacked()
        {
            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelDnaPublishPathNone.csproj /t:Restore,Build /p:Configuration=Release /p:ExcelDnaUnpack=true /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dll", "ExcelDna.Integration.dll", "ExcelDna.Loader.dll", "ExcelDna.ManagedHost.dll", "SDKExcelDnaPublishPathNone.dll");
        }

        [Test]
        public void CleanPacked()
        {
            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelDnaPublishPathNone.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            MsBuild(projectBasePath + "SDKExcelDnaPublishPathNone.csproj /t:Clean /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertNotFound(Path.Combine(projectOutDir, "SDKExcelDnaPublishPathNone-AddIn64-packed.xll"));
            AssertNotFound(Path.Combine(projectOutDir, "SDKExcelDnaPublishPathNone-AddIn-packed.xll"));
        }

        [Test]
        public void CleanUnpacked()
        {
            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelDnaPublishPathNone.csproj /t:Restore,Build /p:Configuration=Release /p:ExcelDnaUnpack=true /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            MsBuild(projectBasePath + "SDKExcelDnaPublishPathNone.csproj /t:Clean /p:Configuration=Release /p:ExcelDnaUnpack=true /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertNotFound(Path.Combine(projectOutDir, "ExcelDna.ManagedHost.dll"));
        }
    }
}
