using System;
using Microsoft.Build.Framework;
using ExcelDna.AddIn.Tasks.Logging;
using ExcelDna.AddIn.Tasks.Utils;
using ExcelDna.PackedResources.Logging;
using System.IO;
using System.Collections.Generic;
using System.Linq;

namespace ExcelDna.AddIn.Tasks
{
    public class PackExcelAddInNativeAOT : AbstractTask
    {
        private readonly IBuildLogger _log;
        private readonly IExcelDnaFileSystem _fileSystem;

        public PackExcelAddInNativeAOT()
        {
            _log = new BuildLogger(this, "PackExcelAddInNativeAOT");
            _fileSystem = new ExcelDnaPhysicalFileSystem();
        }

        internal PackExcelAddInNativeAOT(IBuildLogger log, IExcelDnaFileSystem fileSystem)
        {
            _log = log ?? throw new ArgumentNullException(nameof(log));
            _fileSystem = fileSystem ?? throw new ArgumentNullException(nameof(fileSystem));
        }

        public override bool Execute()
        {
            try
            {
                _log.Debug("Running PackExcelAddInNativeAOT Task");

                bool useManagedResourceResolver = false;
#if NETCOREAPP
                useManagedResourceResolver = PackManagedOnWindows || !OperatingSystem.IsWindows();
#endif

                string mainNativeAssembly = Path.Combine(NativeOutputPath, ProjectName + ".dll");
                IEnumerable<string> includeAssemblies = BuildTaskCommon.SplitDlls(AddInInclude, OutDirectory).Select(i => Path.Combine(OutDirectory, i));
                string xllOutput = Path.Combine(PublishDir, ProjectName + "-AddIn64.xll");
                File.Copy(Xll64FilePath, xllOutput, true);

                int result = ExcelDna.PackedResources.ExcelDnaPack.PackNativeAOT(mainNativeAssembly, includeAssemblies, xllOutput, RunMultithreaded, useManagedResourceResolver, IncludePdb, _log);
                if (result != 0)
                    throw new ApplicationException($"Pack failed with exit code {result}.");

                return true;
            }
            catch (Exception ex)
            {
                _log.Error(ex, ex.Message);
                _log.Error(ex, ex.ToString());
                return false;
            }
        }

        /// <summary>
        /// The 64-bit .xll file path
        /// </summary>
        [Required]
        public string Xll64FilePath { get; set; }

        /// <summary>
        /// The name of the project being compiled
        /// </summary>
        [Required]
        public string ProjectName { get; set; }

        /// <summary>
        /// The output location for the publish target; includes the trailing backslash (\).
        /// </summary>
        [Required]
        public string PublishDir { get; set; }

        /// <summary>
        /// The directory in which the built files were written to
        /// </summary>
        [Required]
        public string OutDirectory { get; set; }

        /// <summary>
        /// The directory in which the native built files were written to
        /// </summary>
        [Required]
        public string NativeOutputPath { get; set; }

        /// <summary>
        /// Use multi threading
        /// </summary>
        [Required]
        public bool RunMultithreaded { get; set; }

        /// <summary>
        /// Enable/disable cross-platform resource packing implementation when executing on Windows.
        /// </summary>
        public bool PackManagedOnWindows { get; set; }

        /// <summary>
        /// Semicolon separated list of references
        /// </summary>
        public string AddInInclude { get; set; }

        /// <summary>
        /// Enable/disable including pdb files in packed add-in
        /// </summary>
        public bool IncludePdb { get; set; }
    }
}
