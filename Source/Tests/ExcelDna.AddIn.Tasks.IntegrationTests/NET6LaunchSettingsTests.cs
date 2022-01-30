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
                DeleteProperties(projectBasePath);
            }
        }

        [Test]
        [TestCase("launchSettingsExistingProfile.json")]
        [TestCase("launchSettingsNoProfile.json")]
        [TestCase("launchSettingsNoProfiles.json")]
        public void UpdateExistingProfile(string src)
        {
            const string projectBasePath = @"NET6LaunchSettings\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            try
            {
                string launchSettingsPath = Path.Combine(projectBasePath, "Properties", "launchSettings.json");
                CopyLaunchSettings(src, launchSettingsPath);

                MsBuild(projectBasePath + "NET6LaunchSettings.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

                AssertFileContains(launchSettingsPath, "NET6LaunchSettings-AddIn");
                AssertFileContains(launchSettingsPath, "67890");
            }
            finally
            {
                DeleteProperties(projectBasePath);
            }
        }

        [Test]
        public void Disabled()
        {
            const string projectBasePath = @"NET6LaunchSettingsDisabled\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            try
            {
                MsBuild(projectBasePath + "NET6LaunchSettingsDisabled.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

                AssertNotFound(Path.Combine(projectBasePath, "Properties", "launchSettings.json"));
            }
            finally
            {
                DeleteProperties(projectBasePath);
            }
        }

        private void DeleteProperties(string projectBasePath)
        {
            string properties = Path.Combine(projectBasePath, "Properties");
            if (Directory.Exists(properties))
                Directory.Delete(properties, true);
        }

        private void CopyLaunchSettings(string srcName, string dstPath)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(dstPath));
            string srcPath = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), srcName);
            File.Copy(srcPath, dstPath, true);
        }
    }
}
