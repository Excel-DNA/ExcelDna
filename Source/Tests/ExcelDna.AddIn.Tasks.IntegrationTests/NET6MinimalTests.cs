using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class NET6MinimalTests : IntegrationTestBase
    {
        [Test]
        public void NET6MinimalTest()
        {
            const string projectBasePath = @"NET6Minimal\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET6Minimal.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));
        }

        [Test]
        public void Compression()
        {
            const string projectBasePath = @"NET6Minimal\";
            const string projectOutDir = projectBasePath + @"bin\Release\";
            string packedFile = Path.Combine(projectOutDir, @"publish\MyLibrary-AddIn64-packed.xll");

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET6Minimal.csproj /t:Restore,Build /p:ExcelDnaPackCompressResources=false /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));
            long notCompressed = (new FileInfo(packedFile)).Length;

            MsBuild(projectBasePath + "NET6Minimal.csproj /t:Restore,Build /p:ExcelDnaPackCompressResources=true /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));
            long compressed = (new FileInfo(packedFile)).Length;

            Assert.Less(compressed, notCompressed - 10000); // 10000 is enough to make sure ExcelDna.Integration and ExcelDna.Loader are not compressed, not only NET6Minimal.dll.
        }

        [Test]
        public void RunMultithreadedDisabled()
        {
            const string projectBasePath = @"NET6Minimal\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET6Minimal.csproj /t:Restore,Build /p:ExcelDnaPackRunMultithreaded=false /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));
        }

        [Test]
        public void RunMultithreadedEnabled()
        {
            const string projectBasePath = @"NET6Minimal\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET6Minimal.csproj /t:Restore,Build /p:ExcelDnaPackRunMultithreaded=true /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));
        }
    }
}
