using System;
using System.IO;
using System.Linq;
using ExcelDna.AddIn.Tasks.Logging;
using Microsoft.Build.Framework;
using ExcelDna.AddIn.Tasks.Utils;

namespace ExcelDna.AddIn.Tasks
{
    public abstract class ExcelDebugTask : AbstractTask
    {
        protected IBuildLogger _log;
        internal Lazy<IExcelDetector> _excelDetectorLazy;
        internal IExcelDetector _excelDetector;

        protected string GetExcelPath()
        {
            var excelExePath = ExcelExePath;

            try
            {
                if (string.IsNullOrWhiteSpace(excelExePath))
                {
                    if (!_excelDetector.TryFindLatestExcel(out excelExePath))
                    {
                        _log.Warning("DNA" + "EXCEL.EXE".GetHashCode(), "Unable to find path where EXCEL.EXE is located");
                        return excelExePath;
                    }
                }

                if (!File.Exists(excelExePath))
                {
                    _log.Warning("DNA" + "EXCEL.EXE".GetHashCode(),
                        "EXCEL.EXE not found on disk at location " + excelExePath);
                }
            }
            finally
            {
                _log.Information("EXCEL.EXE path for debugging: " + excelExePath);
            }

            return excelExePath;
        }

        protected string GetAddInForDebugging(string excelExePath)
        {
            var addInForDebugging = AddInForDebugging;

            try
            {
                if (string.IsNullOrWhiteSpace(addInForDebugging))
                {
                    if (!TryGetExcelAddInForDebugging(excelExePath, out addInForDebugging))
                    {
                        _log.Warning("DNA" + "ADDIN".GetHashCode(), "Unable to find add-in to Debug");
                    }
                }
            }
            finally
            {
                _log.Information("Add-In for debugging: " + addInForDebugging);
            }

            return addInForDebugging;
        }

        private bool TryGetExcelAddInForDebugging(string excelExePath, out string addinForDebugging)
        {
            addinForDebugging = null;

            if (!_excelDetector.TryFindExcelBitness(excelExePath, out var excelBitness))
            {
                return false;
            }

            BuildTaskCommon _common = new BuildTaskCommon(FilesInProject, OutDirectory, FileSuffix32Bit, FileSuffix64Bit, ProjectName, AddInFileName);

            var outputBuildItems = _common.GetBuildItemsForDnaFiles();

            var firstAddIn = outputBuildItems.FirstOrDefault();
            if (firstAddIn == null) return false;

            switch (excelBitness)
            {
                case Bitness.Bit32:
                    {
                        addinForDebugging = firstAddIn.OutputXllFileNameAs32Bit;
                        return true;
                    }
                case Bitness.Bit64:
                    {
                        addinForDebugging = firstAddIn.OutputXllFileNameAs64Bit;
                        return true;
                    }
                default:
                    {
                        return false;
                    }
            }
        }

        /// <summary>
        /// The path to EXCEL.EXE that should be used for debugging
        /// This overrides the automatic detection of the latest Excel installed
        /// </summary>
        public string ExcelExePath { get; set; }

        /// <summary>
        /// The path to .XLL file name that should be used for debugging
        /// This overrides the automatic detection depending on Excel's bitness
        /// </summary>
        public string AddInForDebugging { get; set; }

        /// <summary>
        /// The name of the project being compiled
        /// </summary>
        [Required]
        public string ProjectName { get; set; }

        /// <summary>
        /// The list of files in the project marked as Content or None
        /// </summary>
        [Required]
        public ITaskItem[] FilesInProject { get; set; }

        /// <summary>
        /// The directory in which the built files were written to
        /// </summary>
        [Required]
        public string OutDirectory { get; set; }

        /// <summary>
        /// The name suffix for 32-bit .dna files
        /// </summary>
        public string FileSuffix32Bit { get; set; }

        /// <summary>
        /// The name suffix for 64-bit .dna files
        /// </summary>
        public string FileSuffix64Bit { get; set; }

        /// <summary>
        /// Custom add-in file name
        /// </summary>
        public string AddInFileName { get; set; }
    }
}
