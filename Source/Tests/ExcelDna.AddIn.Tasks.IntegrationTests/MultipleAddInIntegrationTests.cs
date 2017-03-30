using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class MultipleAddInIntegrationTests : IntegrationTestBase
    {
        [Test]
        public void Multiple_AddIn_Projects_Built_To_The_Same_Directory_Should_Only_Clean_Themselves()
        {
            const string projectOneBasePath = @"MultipleAddInProjectOne\";
            const string projectTwoBasePath = @"MultipleAddInProjectTwo\";

            const string projectOutDir = projectOneBasePath + @"..\MultipleAddInBuild\";

            MsBuild(projectOneBasePath + "MultipleAddInProjectOne.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"..\MultipleAddInBuild\"));
            MsBuild(projectTwoBasePath + "MultipleAddInProjectTwo.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"..\MultipleAddInBuild\"));
            MsBuild(projectTwoBasePath + "MultipleAddInProjectTwo.csproj /t:Clean /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"..\MultipleAddInBuild\"));

            //The .DNA files, XLL + config files should remain for project one
            AssertFound(projectOutDir, "*.dna", "AddIn-One-x.dna", "AddIn-One-x64.dna");
            AssertFound(projectOutDir, "*.xll", "AddIn-One-x.xll", "AddIn-One-x-packed.xll", "AddIn-One-x64.xll", "AddIn-One-x64-packed.xll");
            AssertFound(projectOutDir, "*.xll.config", new string[0]);

            AssertIdentical(projectOneBasePath + "AddIn-One-x64.dna", projectOutDir + "AddIn-One-x64.dna");
            AssertIdentical(projectOneBasePath + "AddIn-One-x64.config", projectOutDir + "AddIn-One-x64.xll.config");

            //Assert project two files have been removed
            AssertNotFound(projectOutDir + "AddIn-Two-x-packed.xll");
            AssertNotFound(projectOutDir + "AddIn-Two-x64-packed.xll.config");
            AssertNotFound(projectOutDir + "AddIn-Two-x64-packed.xll");
            AssertNotFound(projectOutDir + "AddIn-Two-x64.xll.config");
            AssertNotFound(projectOutDir + "AddIn-Two-x64.xll");
            AssertNotFound(projectOutDir + "AddIn-Two-x64.dna");
            AssertNotFound(projectOutDir + "AddIn-Two-x.xll");
            AssertNotFound(projectOutDir + "AddIn-Two-x.dna");
        }
    }
}

