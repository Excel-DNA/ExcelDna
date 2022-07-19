using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class ProjectWithNameEndingIn64IntegrationTests : IntegrationTestBase
    {
        [Test]
        public void A_project_with_name_ending_in_64_and_dna_name_ending_in_64_copies_source_dna_to_32_and_64_bit_variants_and_copies_corresponding_xll_files()
        {
            const string projectBasePath = @"ProjectWithNameEndingIn64\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "ProjectWithNameEndingIn64.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dna", "ProjectWithNameEndingIn64-AddIn.dna", "ProjectWithNameEndingIn64-AddIn64.dna");
            AssertOutput(projectOutDir, "*.xll", "ProjectWithNameEndingIn64-AddIn.xll", @"publish\ProjectWithNameEndingIn64-AddIn-packed.xll", "ProjectWithNameEndingIn64-AddIn64.xll", @"publish\ProjectWithNameEndingIn64-AddIn64-packed.xll");
            AssertOutput(projectOutDir, "*.xll.config", "ProjectWithNameEndingIn64-AddIn.xll.config", @"publish\ProjectWithNameEndingIn64-AddIn-packed.xll.config", "ProjectWithNameEndingIn64-AddIn64.xll.config", @"publish\ProjectWithNameEndingIn64-AddIn64-packed.xll.config");

            AssertIdentical(projectBasePath + "ProjectWithNameEndingIn64-AddIn.dna", projectOutDir + "ProjectWithNameEndingIn64-AddIn.dna");
            AssertIdentical(projectBasePath + "ProjectWithNameEndingIn64-AddIn.dna", projectOutDir + "ProjectWithNameEndingIn64-AddIn64.dna");

            Assert32BitXll(projectOutDir + "ProjectWithNameEndingIn64-AddIn.xll");
            Assert64BitXll(projectOutDir + "ProjectWithNameEndingIn64-AddIn64.xll");

            AssertIdentical(projectBasePath + "App.config", projectOutDir + "ProjectWithNameEndingIn64-AddIn.xll.config");
            AssertIdentical(projectBasePath + "App.config", projectOutDir + @"publish\ProjectWithNameEndingIn64-AddIn-packed.xll.config");

            AssertIdentical(projectBasePath + "App.config", projectOutDir + "ProjectWithNameEndingIn64-AddIn64.xll.config");
            AssertIdentical(projectBasePath + "App.config", projectOutDir + @"publish\ProjectWithNameEndingIn64-AddIn64-packed.xll.config");
        }
    }
}
