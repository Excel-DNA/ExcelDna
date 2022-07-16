using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    // ReSharper disable once InconsistentNaming
    public class SingleDnaFileCustomSuffix_x86IntegrationTests : IntegrationTestBase
    {
        [Test]
        public void A_project_with_a_specific_dna_file_for_32_bit_using_custom_x86_suffix_copies_source_file_to_output_and_copies_corresponding_xll_files()
        {
            const string projectBasePath = @"SingleDnaFileCustomSuffix_x86\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SingleDnaFileCustomSuffix_x86.csproj /p:ExcelDna32BitAddInSuffix=-x86 /p:ExcelDna64BitAddInSuffix=-x64 /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dna", "MyLibrary-AddIn-x86.dna", "MyLibrary-AddIn-x64.dna");
            AssertOutput(projectOutDir, "*.xll", "MyLibrary-AddIn-x86.xll", @"publish\MyLibrary-AddIn-x86-packed.xll", "MyLibrary-AddIn-x64.xll", @"publish\MyLibrary-AddIn-x64-packed.xll");
            AssertOutput(projectOutDir, "*.xll.config", "MyLibrary-AddIn-x86.xll.config", @"publish\MyLibrary-AddIn-x86-packed.xll.config", "MyLibrary-AddIn-x64.xll.config", @"publish\MyLibrary-AddIn-x64-packed.xll.config");

            AssertIdentical(projectBasePath + "MyLibrary-AddIn-x86.dna", projectOutDir + "MyLibrary-AddIn-x86.dna");
            AssertIdentical(projectBasePath + "MyLibrary-AddIn-x86.dna", projectOutDir + "MyLibrary-AddIn-x64.dna");

            Assert32BitXll(projectOutDir + "MyLibrary-AddIn-x86.xll");
            Assert64BitXll(projectOutDir + "MyLibrary-AddIn-x64.xll");

            AssertIdentical(projectBasePath + "App.config", projectOutDir + "MyLibrary-AddIn-x86.xll.config");
            AssertIdentical(projectBasePath + "App.config", projectOutDir + @"publish\MyLibrary-AddIn-x86-packed.xll.config");

            AssertIdentical(projectBasePath + "App.config", projectOutDir + "MyLibrary-AddIn-x64.xll.config");
            AssertIdentical(projectBasePath + "App.config", projectOutDir + @"publish\MyLibrary-AddIn-x64-packed.xll.config");
        }
    }
}
