using System;
using System.Collections;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Build.Framework;
using ExcelDna.AddIn.Tasks.Utils;
using Microsoft.Build.Utilities;

namespace ExcelDna.AddIn.Tasks
{
    public class CleanExcelAddIn : AbstractTask
    {
        private readonly IExcelDnaFileSystem _fileSystem;
        private List<ITaskItem> _packedFilesToDelete;
        private BuildTaskCommon _common;

        public CleanExcelAddIn()
            : this(new ExcelDnaPhysicalFileSystem())
        {
        }

        public CleanExcelAddIn(IExcelDnaFileSystem fileSystem)
            : base("ExcelDnaClean")
        {
            if (fileSystem == null)
            {
                throw new ArgumentNullException("fileSystem");
            }

            _fileSystem = fileSystem;
        }

        public override bool Execute()
        {
            try
            {
                LogDebugMessage("Running CleanExcelAddIn MSBuild Task");

                LogDiagnostics();

                FilesInProject = FilesInProject ?? new ITaskItem[0];
                LogDebugMessage("Number of files in project: " + FilesInProject.Length);

                _common = new BuildTaskCommon(FilesInProject, OutDirectory, FileSuffix32Bit, FileSuffix64Bit);

                var existingBuiltFiles = _common.GetBuildItemsForDnaFiles();
                _packedFilesToDelete = GetPackedFilesToDelete(existingBuiltFiles);

                // Get the packed name versions : Refactor this + build items
                DeleteAddInFiles(existingBuiltFiles);
                DeletePackedAddInFiles(_packedFilesToDelete);

                return true;
            }
            catch (Exception ex)
            {
                LogError("DNA" + ex.GetType().Name.GetHashCode(), ex.Message);
                LogError("DNA" + ex.GetType().Name.GetHashCode(), ex.ToString());
                return false;
            }
        }

        private void LogDiagnostics()
        {
            LogDebugMessage("----Arguments----");
            LogDebugMessage("FilesInProject: " + (FilesInProject ?? new ITaskItem[0]).Length);
            LogDebugMessage("OutDirectory: " + OutDirectory);
            LogDebugMessage("Xll32FilePath: " + Xll32FilePath);
            LogDebugMessage("Xll64FilePath: " + Xll64FilePath);
            LogDebugMessage("FileSuffix32Bit: " + FileSuffix32Bit);
            LogDebugMessage("FileSuffix64Bit: " + FileSuffix64Bit);
            LogDebugMessage("-----------------");
        }

        private List<ITaskItem> GetPackedFilesToDelete(BuildItemSpec[] existingBuiltFiles)
        {
            var packedFilesToDelete = new List<ITaskItem>();

            foreach (var item in existingBuiltFiles)
            {
                packedFilesToDelete.Add(GetPackedFileNames(item.OutputDnaFileNameAs32Bit, item.OutputXllFileNameAs32Bit, item.OutputConfigFileNameAs32Bit));
                packedFilesToDelete.Add(GetPackedFileNames(item.OutputDnaFileNameAs64Bit, item.OutputXllFileNameAs64Bit, item.OutputConfigFileNameAs64Bit));
            }

            return packedFilesToDelete;
        }

        private TaskItem GetPackedFileNames(string outputDnaFileName, string outputXllFileName, string outputXllConfigFileName)
        {
            var outputPackedDnaFileName = !string.IsNullOrWhiteSpace(PackedFileSuffix)
            ? Path.Combine(Path.GetDirectoryName(outputDnaFileName) ?? string.Empty,
                Path.GetFileNameWithoutExtension(outputDnaFileName) + PackedFileSuffix + ".dna")
            : outputDnaFileName;

            var outputPackedXllFileName = !string.IsNullOrWhiteSpace(PackedFileSuffix)
              ? Path.Combine(Path.GetDirectoryName(outputXllFileName) ?? string.Empty,
                  Path.GetFileNameWithoutExtension(outputXllFileName) + PackedFileSuffix + ".xll")
              : outputXllFileName;

            var outputPackedXllConfigFileName = !string.IsNullOrWhiteSpace(PackedFileSuffix)
            ? Path.Combine(Path.GetDirectoryName(outputXllFileName) ?? string.Empty,
                Path.GetFileNameWithoutExtension(outputXllFileName) + PackedFileSuffix + ".xll.config")
            : outputXllConfigFileName;

            var metadata = new Hashtable
            {
                {"OutputPackedDnaFileName", outputPackedDnaFileName},
                {"OutputPackedXllFileName", outputPackedXllFileName},
                {"OutputPackedXllConfigFileName", outputPackedXllConfigFileName },
            };

            return new TaskItem(outputDnaFileName, metadata);
        }

        private void DeleteAddInFiles(BuildItemSpec[] buildItemsForDnaFiles)
        {
            foreach (var item in buildItemsForDnaFiles)
            {
                DeleteFileIfExists(item.OutputDnaFileNameAs32Bit);
                DeleteFileIfExists(item.OutputDnaFileNameAs64Bit);

                DeleteFileIfExists(item.OutputXllFileNameAs32Bit);
                DeleteFileIfExists(item.OutputXllFileNameAs64Bit);

                DeleteFileIfExists(item.OutputConfigFileNameAs32Bit);
                DeleteFileIfExists(item.OutputConfigFileNameAs64Bit);
            }
        }

        private void DeletePackedAddInFiles(List<ITaskItem> filesToDelete)
        {
            filesToDelete.ToList().ForEach(f =>
            {
                DeleteFileIfExists(f.GetMetadata("OutputPackedDnaFileName"));
                DeleteFileIfExists(f.GetMetadata("OutputPackedXllFileName"));
                DeleteFileIfExists(f.GetMetadata("OutputPackedXllConfigFileName"));
            });
        }

        private void DeleteFileIfExists(string path)
        {
            if (_fileSystem.FileExists(path))
            {
                LogDebugMessage("Deleting file " + path);
                _fileSystem.DeleteFile(path);
            }
        }

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
        /// The 32-bit .xll file path; set to <code>$(MSBuildThisFileDirectory)\ExcelDna.xll</code> by default
        /// </summary>
        [Required]
        public string Xll32FilePath { get; set; }

        /// <summary>
        /// The 64-bit .xll file path; set to <code>$(MSBuildThisFileDirectory)\ExcelDna64.xll</code> by default
        /// </summary>
        [Required]
        public string Xll64FilePath { get; set; }

        /// <summary>
        /// The name suffix for 32-bit .dna files
        /// </summary>
        public string FileSuffix32Bit { get; set; }

        /// <summary>
        /// The name suffix for 64-bit .dna files
        /// </summary>
        public string FileSuffix64Bit { get; set; }

        /// <summary>
        /// Enable/disable running ExcelDnaPack for .dna files
        /// </summary>
        public bool PackIsEnabled { get; set; }

        /// <summary>
        /// Enable/disable running ExcelDnaPack for .dna files
        /// </summary>
        public string PackedFileSuffix { get; set; }
    }
}
