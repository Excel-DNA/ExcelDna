using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SeparateDnaFilesDefaultSuffixIntegrationTests : IntegrationTestBase
    {
        [Test]
        public void A_project_with_a_specific_dna_files_for_32_and_64_bit_using_default_suffix_copies_each_source_file_to_output_and_copies_corresponding_xll_files()
        {
            const string projectBasePath = @"SeparateDnaFilesDefaultSuffix\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SeparateDnaFilesDefaultSuffix.csproj /t:Build /p:Configuration=Release /p:ExcelDnaPublishPath=publish /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dna", "MyLibrary-AddIn.dna", "MyLibrary-AddIn64.dna");
            AssertOutput(projectOutDir, "*.xll", "MyLibrary-AddIn.xll", @"publish\MyLibrary-AddIn-packed.xll", "MyLibrary-AddIn64.xll", @"publish\MyLibrary-AddIn64-packed.xll");
            AssertOutput(projectOutDir, "*.xll.config", "MyLibrary-AddIn.xll.config", @"publish\MyLibrary-AddIn-packed.xll.config", "MyLibrary-AddIn64.xll.config", @"publish\MyLibrary-AddIn64-packed.xll.config");

            AssertIdentical(projectBasePath + "MyLibrary-AddIn.dna", projectOutDir + "MyLibrary-AddIn.dna");
            AssertIdentical(projectBasePath + "MyLibrary-AddIn64.dna", projectOutDir + "MyLibrary-AddIn64.dna");

            Assert32BitXll(projectOutDir + "MyLibrary-AddIn.xll");
            Assert64BitXll(projectOutDir + "MyLibrary-AddIn64.xll");

            AssertIdentical(projectBasePath + "App.config", projectOutDir + "MyLibrary-AddIn.xll.config");
            AssertIdentical(projectBasePath + "App.config", projectOutDir + @"publish\MyLibrary-AddIn-packed.xll.config");

            AssertIdentical(projectBasePath + "App64.config", projectOutDir + "MyLibrary-AddIn64.xll.config");
            AssertIdentical(projectBasePath + "App64.config", projectOutDir + @"publish\MyLibrary-AddIn64-packed.xll.config");
        }
    }
}
