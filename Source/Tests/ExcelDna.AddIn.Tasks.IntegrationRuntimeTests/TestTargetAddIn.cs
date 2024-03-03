using System.IO;
using ExcelDna.Testing;

namespace ExcelDna.AddIn.Tasks.IntegrationRuntimeTests
{
    internal class TestTargetAddIn
    {
        public static void Build(string projectName, string? parameters = null)
        {
            Directory.Delete(Path.Combine(GetProjectDir(projectName), outputPath), true);

            MsBuild.Build(Path.Combine(GetProjectDir(projectName), $"{projectName}.csproj") + $" /t:Restore,Build /p:Configuration=Release {parameters} /v:m " + MsBuild.Param("OutputPath", outputPath));
        }

        public static bool Register(string projectName, string xllName)
        {
            string xllPath = Path.Combine(GetProjectDir(projectName), outputPath, xllName);

            return Util.Application.RegisterXLL(xllPath);
        }

        private static string GetProjectDir(string projectName)
        {
            string testTargetPath = Path.GetFullPath(@"..\..\..\..\ExcelDna.AddIn.Tasks.IntegrationTests.TestTarget");
            return Path.Combine(testTargetPath, projectName);
        }

        private const string outputPath = @"bin\Release\";
    }
}
