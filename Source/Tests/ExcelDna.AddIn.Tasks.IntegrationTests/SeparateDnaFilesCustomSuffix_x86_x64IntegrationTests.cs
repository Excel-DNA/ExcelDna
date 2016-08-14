using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    // ReSharper disable once InconsistentNaming
    public class SeparateDnaFilesCustomSuffix_x86_x64IntegrationTests : IntegrationTestBase
    {
        [Test]
        public void A_project_with_specific_dna_files_for_32_and_64_bit_using_custom_suffixes_copies_each_source_file_to_output_and_copies_corresponding_xll_files()
        {
            const string projectBasePath = @"SeparateDnaFilesCustomSuffix_x86_x64\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SeparateDnaFilesCustomSuffix_x86_x64.csproj /p:ExcelDna32BitAddInSuffix=-x86 /p:ExcelDna64BitAddInSuffix=-x64 /p:ExcelDnaPackXllSuffix=-bundled /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dna", "MyLibrary-AddIn-x86.dna", "MyLibrary-AddIn-x64.dna");
            AssertOutput(projectOutDir, "*.xll", "MyLibrary-AddIn-x86.xll", "MyLibrary-AddIn-x86-bundled.xll", "MyLibrary-AddIn-x64.xll", "MyLibrary-AddIn-x64-bundled.xll");
            AssertOutput(projectOutDir, "*.xll.config", "MyLibrary-AddIn-x86.xll.config", "MyLibrary-AddIn-x86-bundled.xll.config", "MyLibrary-AddIn-x64.xll.config", "MyLibrary-AddIn-x64-bundled.xll.config");

            AssertIdentical(projectBasePath + "MyLibrary-AddIn-x86.dna", projectOutDir + "MyLibrary-AddIn-x86.dna");
            AssertIdentical(projectBasePath + "MyLibrary-AddIn-x64.dna", projectOutDir + "MyLibrary-AddIn-x64.dna");

            Assert32BitXll(projectOutDir + "MyLibrary-AddIn-x86.xll");
            Assert64BitXll(projectOutDir + "MyLibrary-AddIn-x64.xll");

            AssertIdentical(projectBasePath + "App-x86.config", projectOutDir + "MyLibrary-AddIn-x86.xll.config");
            AssertIdentical(projectBasePath + "App-x86.config", projectOutDir + "MyLibrary-AddIn-x86-bundled.xll.config");

            AssertIdentical(projectBasePath + "App-x64.config", projectOutDir + "MyLibrary-AddIn-x64.xll.config");
            AssertIdentical(projectBasePath + "App-x64.config", projectOutDir + "MyLibrary-AddIn-x64-bundled.xll.config");
        }
    }
}