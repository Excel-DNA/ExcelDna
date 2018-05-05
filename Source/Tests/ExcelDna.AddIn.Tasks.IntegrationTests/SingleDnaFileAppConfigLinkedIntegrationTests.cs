using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SingleDnaFileAppConfigLinkedIntegrationTests : IntegrationTestBase
    {
        [Test]
        public void A_project_with_an_app_config_added_as_link_will_generate_app_config_in_the_release_folder()
        {
            const string appConfigProjectBasePath = @"SingleDnaFileDefaultSuffix\";
            const string projectBasePath = @"SingleDnaFileAppConfigLinked\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SingleDnaFileAppConfigLinked.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertIdentical(appConfigProjectBasePath + "App.config", projectOutDir + "MyLibrary-AddIn.xll.config");
            AssertIdentical(appConfigProjectBasePath + "App.config", projectOutDir + "MyLibrary-AddIn-packed.xll.config");

            AssertIdentical(appConfigProjectBasePath + "App.config", projectOutDir + "MyLibrary-AddIn64.xll.config");
            AssertIdentical(appConfigProjectBasePath + "App.config", projectOutDir + "MyLibrary-AddIn64-packed.xll.config");
        }
    }
}
