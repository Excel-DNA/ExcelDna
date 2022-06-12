using System;
using System.Collections;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;
using ExcelDna.AddIn.Tasks.Logging;
using ExcelDna.AddIn.Tasks.Utils;

namespace ExcelDna.AddIn.Tasks
{
    public class PackExcelAddIn : AbstractTask
    {
        private readonly IBuildLogger _log;
        private readonly IExcelDnaFileSystem _fileSystem;

        public PackExcelAddIn()
        {
            _log = new BuildLogger(this, "PackExcelAddIn");
            _fileSystem = new ExcelDnaPhysicalFileSystem();
        }

        internal PackExcelAddIn(IBuildLogger log, IExcelDnaFileSystem fileSystem)
        {
            _log = log ?? throw new ArgumentNullException(nameof(log));
            _fileSystem = fileSystem ?? throw new ArgumentNullException(nameof(fileSystem));
        }

        public override bool Execute()
        {
            try
            {
                _log.Debug("Running PackExcelAddIn Task");

                return ExcelDna.PackedResources.ExcelDnaPack.Pack(OutputDnaFileName, OutputPackedXllFileName, CompressResources, RunMultithreaded, true, null) == 0;
            }
            catch (Exception ex)
            {
                _log.Error(ex, ex.Message);
                _log.Error(ex, ex.ToString());
                return false;
            }
        }

        /// <summary>
        /// The path to the primary .dna file for the ExcelDna add-in
        /// </summary>
        [Required]
        public string OutputDnaFileName { get; set; }

        /// <summary>
        /// Output path
        /// </summary>
        [Required]
        public string OutputPackedXllFileName { get; set; }

        /// <summary>
        /// Compress (LZMA) of resources
        /// </summary>
        [Required]
        public bool CompressResources { get; set; }

        /// <summary>
        /// Use multi threading
        /// </summary>
        [Required]
        public bool RunMultithreaded { get; set; }
    }
}
