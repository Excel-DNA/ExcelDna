using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using ExcelDna.AddIn.Tasks.IntegrationTests.Utils;
using Microsoft.VisualStudio.Setup.Configuration;
using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    public abstract class IntegrationTestBase
    {
        private string _originalDirectory;

        [SetUp]
        public virtual void SetUp()
        {
            _originalDirectory = Environment.CurrentDirectory;

            var thisAssemblyPath = new Uri(typeof(IntegrationTestBase).Assembly.CodeBase).LocalPath;
            string testsDirName = @"\Source\Tests\";
            string testsDirPath = thisAssemblyPath.Substring(0, thisAssemblyPath.LastIndexOf(testsDirName) + testsDirName.Length);
            Environment.CurrentDirectory = Path.Combine(testsDirPath, @"ExcelDna.AddIn.Tasks.IntegrationTests.TestTarget");
        }

        [TearDown]
        public virtual void TearDown()
        {
            Environment.CurrentDirectory = _originalDirectory;
        }

        protected static void MsBuild(string commandLineArguments)
        {
            MsBuild(commandLineArguments, null);
        }

        protected static void MsBuild(string commandLineArguments, Action<string> outputValidator)
        {
            var allOutput = new StringBuilder();

            Action<string> writer = output =>
            {
                allOutput.AppendLine(output);
                Console.WriteLine(output);
            };

            var msBuildPath = GetMsBuildPath();

            var result = SilentProcessRunner.ExecuteCommand(msBuildPath, commandLineArguments, Environment.CurrentDirectory, writer, e => writer("ERROR: " + e));
            if (result != 0)
            {
                Assert.Fail("MSBuild returned a non-zero exit code: " + result);
            }

            if (outputValidator != null)
            {
                outputValidator(allOutput.ToString());
            }
        }

        protected string MsBuildParam(string name, string value)
        {
            return string.Format("/p:{0}=\"{1}\\\"", name, value);
        }

        private static string GetMsBuildPath()
        {
            try
            {
                var vsConfiguration = new SetupConfiguration();
                var vsInstancesEnumerator = vsConfiguration.EnumAllInstances();
                int fetched;
                var vsInstances = new ISetupInstance[1];
                do
                {
                    vsInstancesEnumerator.Next(1, vsInstances, out fetched);
                    if (fetched > 0)
                    {
                        var vsInstance = vsInstances[0];
                        if (vsInstance.GetInstallationVersion().StartsWith("17."))
                            return vsInstance.ResolvePath(@"Msbuild\Current\Bin\amd64\MSBuild.exe");
                    }
                }
                while (fetched > 0);
            }
            catch (Exception)
            {
            }

            string msBuildPath;

            var programFilesDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            foreach (var version in new[] { "14.0", "12.0" })
            {
                var buildDirectory = Path.Combine(programFilesDirectory, "MSBuild", version, "Bin");
                msBuildPath = Path.Combine(buildDirectory, "msbuild.exe");

                if (File.Exists(msBuildPath))
                {
                    return msBuildPath;
                }
            }

            var netFx = System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory();

            msBuildPath = Path.Combine(netFx, "msbuild.exe");
            if (!File.Exists(msBuildPath))
            {
                Assert.Fail("Could not find MSBuild at: " + msBuildPath);
            }

            return msBuildPath;
        }

        protected static void Clean(string path)
        {
            if (!Directory.Exists(path))
            {
                return;
            }

            var filesToDelete = Directory.GetFiles(path, "*.*", SearchOption.AllDirectories);
            foreach (var file in filesToDelete)
            {
                File.Delete(file);
            }
        }

        protected static void AssertOutput(string basePath, string searchPattern, params string[] expectedFiles)
        {
            var existingFiles = Directory.GetFiles(basePath, searchPattern, SearchOption.AllDirectories)
                .ToArray();

            var expectedFilesWithBasePath = expectedFiles
                .Select(f => Path.Combine(basePath, f))
                .ToArray();

            var filesMissing = expectedFilesWithBasePath.Except(existingFiles, StringComparer.OrdinalIgnoreCase).ToArray();
            if (filesMissing.Length > 0)
            {
                Assert.Fail("Expected file(s) missing in the output: {0}", string.Join(", ", filesMissing));
            }

            var extraFiles = existingFiles.Except(expectedFilesWithBasePath, StringComparer.OrdinalIgnoreCase).ToArray();
            if (extraFiles.Length > 0)
            {
                Assert.Fail("Unexpected file(s) found in the output: {0}", string.Join(", ", extraFiles));
            }
        }

        protected static void AssertFound(string basePath, string searchPattern, params string[] expectedFiles)
        {
            var existingFiles = Directory.GetFiles(basePath, searchPattern, SearchOption.AllDirectories)
               .ToArray();

            var expectedFilesWithBasePath = expectedFiles
                .Select(f => Path.Combine(basePath, f))
                .ToArray();

            var filesMissing = expectedFilesWithBasePath.Except(existingFiles, StringComparer.OrdinalIgnoreCase).ToArray();
            if (filesMissing.Length > 0)
            {
                Assert.Fail("Expected file(s) missing in the output: {0}", string.Join(", ", filesMissing));
            }
        }

        protected void AssertNotFound(string fileName)
        {
            if (File.Exists(fileName))
            {
                Assert.Fail("File {0} exists", fileName);
            }
        }

        protected void AssertIdentical(string fileName1, string fileName2)
        {
            if (!File.Exists(fileName1))
            {
                Assert.Fail("File {0} does not exist", fileName1);
            }

            if (!File.Exists(fileName2))
            {
                Assert.Fail("File {0} does not exist", fileName2);
            }

            Assert.IsTrue(FilesHaveEqualHash(fileName1, fileName2), "Contents of {0} and {1} do not match",
                fileName1, fileName2);
        }

        protected void Assert32BitXll(string xllFileName)
        {
            if (!File.Exists(xllFileName))
            {
                Assert.Fail("File {0} does not exist", xllFileName);
            }

            Assert.IsTrue(FilesHaveEqualHash(xllFileName, @"..\.exceldna.addin\tools\ExcelDna.xll"), "{0} is not a 32-bit .xll file",
                xllFileName);
        }

        protected void Assert64BitXll(string xllFileName)
        {
            if (!File.Exists(xllFileName))
            {
                Assert.Fail("File {0} does not exist", xllFileName);
            }

            Assert.IsTrue(FilesHaveEqualHash(xllFileName, @"..\.exceldna.addin\tools\ExcelDna64.xll"), "{0} is not a 64-bit .xll file",
                xllFileName);
        }

        private static bool FilesHaveEqualHash(string file1, string file2)
        {
            var hashFile1 = ComputeFileHash(file1);
            var hashFile2 = ComputeFileHash(file2);

            if (hashFile1.Length != hashFile2.Length)
            {
                return false;
            }

            return !hashFile1.Where((t, i) => t != hashFile2[i]).Any();
        }

        private static byte[] ComputeFileHash(string fileName)
        {
            using (var stream = File.OpenRead(fileName))
            {
                var fileHash = MD5.Create().ComputeHash(stream);
                return fileHash;
            }
        }
    }
}
