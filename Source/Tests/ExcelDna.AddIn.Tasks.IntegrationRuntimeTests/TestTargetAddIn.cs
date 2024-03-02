using System.IO;
using ExcelDna.Testing;

namespace ExcelDna.AddIn.Tasks.IntegrationRuntimeTests
{
    internal class TestTargetAddIn
    {
        public static void Build(string projectName)
        {
            MsBuild.Build(Path.Combine(GetProjectDir(projectName), $"{projectName}.csproj") + " /t:Restore,Build /p:Configuration=Release /v:m " + MsBuild.Param("OutputPath", outputPath));
        }

        public static void Register(string projectName, bool packed)
        {
            string xllPath = Path.Combine(GetProjectDir(projectName), outputPath, packed ? Path.Combine("publish", $"{projectName}-AddIn64-packed.xll") : $"{projectName}-AddIn64.xll");

            Util.Application.RegisterXLL(xllPath);
        }

        private static string GetProjectDir(string projectName)
        {
            string testTargetPath = Path.GetFullPath(@"..\..\..\..\ExcelDna.AddIn.Tasks.IntegrationTests.TestTarget");
            return Path.Combine(testTargetPath, projectName);
        }

        private const string outputPath = @"bin\Release\";
    }
}
