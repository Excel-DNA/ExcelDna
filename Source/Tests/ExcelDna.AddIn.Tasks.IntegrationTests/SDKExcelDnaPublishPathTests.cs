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

        [Test]
        public void Unpacked()
        {
            const string projectBasePath = @"SDKExcelDnaPublishPath\";
            const string projectOutDir = projectBasePath + @"bin\Release\";
            const string publishDir = projectOutDir + @"MyPublish\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelDnaPublishPath.csproj /t:Restore,Build /p:Configuration=Release /p:ExcelDnaUnpack=true /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(publishDir, "*.dll", "ExcelDna.Integration.dll", "ExcelDna.Loader.dll", "ExcelDna.ManagedHost.dll", "SDKExcelDnaPublishPath.dll");
            AssertOutput(publishDir, "*.xll", "SDKExcelDnaPublishPath-AddIn.xll", "SDKExcelDnaPublishPath-AddIn64.xll");
            AssertOutput(publishDir, "*.dna", "SDKExcelDnaPublishPath-AddIn.dna", "SDKExcelDnaPublishPath-AddIn64.dna");
        }
    }
}
