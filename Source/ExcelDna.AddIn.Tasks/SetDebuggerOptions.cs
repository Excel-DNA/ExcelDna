using System;
using System.IO;
using System.Linq;
using Microsoft.Build.Framework;
using ExcelDna.AddIn.Tasks.Utils;

namespace ExcelDna.AddIn.Tasks
{
    public class SetDebuggerOptions : AbstractTask
    {
        private readonly IExcelDetector _excelDetector;
        private readonly IExcelDnaProject _dte;
        private BuildTaskCommon _common;

        public SetDebuggerOptions()
            : this(new ExcelDetector(), new ExcelDnaProject())
        {
        }

        public SetDebuggerOptions(IExcelDetector excelDetector, IExcelDnaProject dte)
            : base("ExcelDnaSetDebuggerOptions")
        {
            if (excelDetector == null)
            {
                throw new ArgumentNullException("excelDetector");
            }

            if (dte == null)
            {
                throw new ArgumentNullException("dte");
            }

            _excelDetector = excelDetector;
            _dte = dte;
        }

        public override bool Execute()
        {
            try
            {
                LogDebugMessage("Running SetDebuggerOptions MSBuild Task");

                LogDiagnostics();

                FilesInProject = FilesInProject ?? new ITaskItem[0];
                LogDebugMessage("Number of files in project: " + FilesInProject.Length);

                var excelExePath = GetExcelPath();
                var addInForDebugging = GetAddInForDebugging(excelExePath);

                if (!_dte.TrySetDebuggerOptions(ProjectName, excelExePath, addInForDebugging))
                {
                    LogWarning("DNA" + "DTE".GetHashCode(), "Unable to set the debugger options within Visual Studio.");
                }
            }
            catch (Exception ex)
            {
                LogWarning("DNA" + ex.GetType().Name.GetHashCode(), ex.Message);
            }

            // Setting the debugger options is not essential to the build process, thus if anything
            // goes wrong, we'll report errors and warnings, but will let the build continue
            return true;
        }

        private string GetExcelPath()
        {
            var excelExePath = ExcelExePath;

            try
            {
                if (string.IsNullOrWhiteSpace(excelExePath))
                {
                    if (!_excelDetector.TryFindLatestExcel(out excelExePath))
                    {
                        LogWarning("DNA" + "EXCEL.EXE".GetHashCode(), "Unable to find path where EXCEL.EXE is located");
                        return excelExePath;
                    }
                }

                if (!File.Exists(excelExePath))
                {
                    LogWarning("DNA" + "EXCEL.EXE".GetHashCode(),
                        "EXCEL.EXE not found on disk at location " + excelExePath);
                }
            }
            finally
            {
                LogMessage("EXCEL.EXE path for debugging: " + excelExePath);
            }

            return excelExePath;
        }

        private string GetAddInForDebugging(string excelExePath)
        {
            var addInForDebugging = AddInForDebugging;

            try
            {
                if (string.IsNullOrWhiteSpace(addInForDebugging))
                {
                    if (!TryGetExcelAddInForDebugging(excelExePath, out addInForDebugging))
                    {
                        LogWarning("DNA" + "ADDIN".GetHashCode(), "Unable to find add-in to Debug");
                    }
                }
            }
            finally
            {
                LogMessage("Add-In for debugging: " + addInForDebugging);
            }

            return addInForDebugging;
        }

        private bool TryGetExcelAddInForDebugging(string excelExePath, out string addinForDebugging)
        {
            addinForDebugging = null;

            Bitness excelBitness;
            if (!_excelDetector.TryFindExcelBitness(excelExePath, out excelBitness))
            {
                return false;
            }

            _common = new BuildTaskCommon(FilesInProject, OutDirectory, FileSuffix32Bit, FileSuffix64Bit);

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
        /// The name of the project being compiled
        /// </summary>
        [Required]
        public string ProjectName { get; set; }

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

        private void LogDiagnostics()
        {
            LogDebugMessage("----Arguments----");
            LogDebugMessage("ProjectName: " + ProjectName);
            LogDebugMessage("ExcelExePath: " + ExcelExePath);
            LogDebugMessage("AddInForDebugging: " + AddInForDebugging);
            LogDebugMessage("FilesInProject: " + (FilesInProject ?? new ITaskItem[0]).Length);
            LogDebugMessage("OutDirectory: " + OutDirectory);
            LogDebugMessage("FileSuffix32Bit: " + FileSuffix32Bit);
            LogDebugMessage("FileSuffix64Bit: " + FileSuffix64Bit);
            LogDebugMessage("-----------------");
        }
    }
}
