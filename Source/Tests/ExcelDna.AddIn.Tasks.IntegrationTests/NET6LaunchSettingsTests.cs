using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class NET6LaunchSettingsTests : IntegrationTestBase
    {
        [Test]
        public void Create()
        {
            const string projectBasePath = @"NET6LaunchSettings\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            try
            {
                MsBuild(projectBasePath + "NET6LaunchSettings.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

                AssertFileContains(Path.Combine(projectBasePath, "Properties", "launchSettings.json"), "NET6LaunchSettings-AddIn");
            }
            finally
            {
                Directory.Delete(Path.Combine(projectBasePath, "Properties"), true);
            }
        }
    }
}
