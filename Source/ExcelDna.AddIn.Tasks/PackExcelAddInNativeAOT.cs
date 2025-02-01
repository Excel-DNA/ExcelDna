using System;
using Microsoft.Build.Framework;
using ExcelDna.AddIn.Tasks.Logging;
using ExcelDna.AddIn.Tasks.Utils;
using ExcelDna.PackedResources.Logging;
using System.IO;

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

                int result = ExcelDna.PackedResources.ExcelDnaPack.PackNativeAOT(OutputPackedXllFileName, RunMultithreaded, useManagedResourceResolver, _log);
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
        /// Output path
        /// </summary>
        [Required]
        public string OutputPackedXllFileName { get; set; }

        /// <summary>
        /// Use multi threading
        /// </summary>
        [Required]
        public bool RunMultithreaded { get; set; }

        /// <summary>
        /// Enable/disable cross-platform resource packing implementation when executing on Windows.
        /// </summary>
        public bool PackManagedOnWindows { get; set; }
    }
}
