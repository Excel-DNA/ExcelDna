using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class NET6LoadFromBytesTests : IntegrationTestBase
    {
        [Test]
        public void Default()
        {
            const string projectBasePath = @"NET6LoadFromBytes\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET6LoadFromBytes.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertFileContains(Path.Combine(projectOutDir, "NET6LoadFromBytes-AddIn64.dna"), "LoadFromBytes=\"true\"");
        }

        [Test]
        public void Disabled()
        {
            const string projectBasePath = @"NET6LoadFromBytes\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET6LoadFromBytes.csproj /t:Restore,Build /p:ExcelAddInLoadFromBytes=false /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertFileContains(Path.Combine(projectOutDir, "NET6LoadFromBytes-AddIn64.dna"), "LoadFromBytes=\"false\"");
        }

        [Test]
        public void Enabled()
        {
            const string projectBasePath = @"NET6LoadFromBytes\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET6LoadFromBytes.csproj /t:Restore,Build /p:ExcelAddInLoadFromBytes=true /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertFileContains(Path.Combine(projectOutDir, "NET6LoadFromBytes-AddIn64.dna"), "LoadFromBytes=\"true\"");
        }
    }
}
