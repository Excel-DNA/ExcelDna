using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class NET6CopyDepsJsonTests : IntegrationTestBase
    {
        [Test]
        public void Copy()
        {
            const string projectBasePath = @"NET6CopyDepsJson\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET6CopyDepsJson.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertFound(projectOutDir, "*.deps.json", new string[] { "NET6CopyDepsJson.deps.json", "NET6CopyDepsJson-AddIn.deps.json", "NET6CopyDepsJson-AddIn64.deps.json" });
        }
    }
}
