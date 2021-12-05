using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class NET5MinimalTests : IntegrationTestBase
    {
        [Test]
        public void NET5MinimalTest()
        {
            const string projectBasePath = @"NET5Minimal\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET5Minimal.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));
        }
    }
}
