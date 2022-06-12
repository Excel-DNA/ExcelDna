using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class NET6UnpackTests : IntegrationTestBase
    {
        [Test]
        public void Unpack()
        {
            const string projectBasePath = @"NET6Unpack\";
            const string projectOutDir = projectBasePath + @"bin\Release";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET6Unpack.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertFound(projectOutDir, "*.dll", new string[] { "ExcelDna.ManagedHost.dll", "ExcelDna.Integration.dll", "ExcelDna.Loader.dll" });
        }

        [Test]
        public void Clean()
        {
            const string projectBasePath = @"NET6Unpack\";
            const string projectOutDir = projectBasePath + @"bin\Release";

            MsBuild(projectBasePath + "NET6Unpack.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            MsBuild(projectBasePath + "NET6Unpack.csproj /t:Clean /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertNotFound(Path.Combine(projectOutDir, "ExcelDna.ManagedHost.dll"));
        }

        [Test]
        public void DisablePack()
        {
            const string projectBasePath = @"NET6Unpack\";
            const string projectOutDir = projectBasePath + @"bin\Release";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET6Unpack.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertNotFound(Path.Combine(projectOutDir, "NET6Unpack-AddIn64-packed.xll"));
        }
    }
}
