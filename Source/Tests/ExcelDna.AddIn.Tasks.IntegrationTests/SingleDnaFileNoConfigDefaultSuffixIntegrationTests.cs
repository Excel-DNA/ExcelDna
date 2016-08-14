using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SingleDnaFileNoConfigDefaultSuffixIntegrationTests : IntegrationTestBase
    {
        [Test]
        public void A_project_with_a_single_dna_file_and_no_config_file_using_default_suffix_copies_source_dna_to_32_and_64_bit_variants_and_copies_corresponding_xll_files()
        {
            const string projectBasePath = @"SingleDnaFileNoConfigDefaultSuffix\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SingleDnaFileNoConfigDefaultSuffix.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dna", "MyLibrary-AddIn.dna", "MyLibrary-AddIn64.dna");
            AssertOutput(projectOutDir, "*.xll", "MyLibrary-AddIn.xll", "MyLibrary-AddIn-packed.xll", "MyLibrary-AddIn64.xll", "MyLibrary-AddIn64-packed.xll");
            AssertOutput(projectOutDir, "*.xll.config", new string[0]);

            AssertIdentical(projectBasePath + "MyLibrary-AddIn.dna", projectOutDir + "MyLibrary-AddIn.dna");
            AssertIdentical(projectBasePath + "MyLibrary-AddIn.dna", projectOutDir + "MyLibrary-AddIn64.dna");

            Assert32BitXll(projectOutDir + "MyLibrary-AddIn.xll");
            Assert64BitXll(projectOutDir + "MyLibrary-AddIn64.xll");
        }
    }
}
