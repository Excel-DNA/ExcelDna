using System;
using System.Collections;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;
using ExcelDna.AddIn.Tasks.Logging;
using ExcelDna.AddIn.Tasks.Utils;
using ExcelDna.PackedResources.Logging;

namespace ExcelDna.AddIn.Tasks
{
    public class CreateExcelAddIn : AbstractTask
    {
        private readonly IBuildLogger _log;
        private readonly IExcelDnaFileSystem _fileSystem;
        private ITaskItem[] _configFilesInProject;
        private List<ITaskItem> _dnaFilesToPack;
        private BuildTaskCommon _common;

        public CreateExcelAddIn()
        {
            _log = new BuildLogger(this, "ExcelDnaBuild");
            _fileSystem = new ExcelDnaPhysicalFileSystem();
        }

        internal CreateExcelAddIn(IBuildLogger log, IExcelDnaFileSystem fileSystem)
        {
            _log = log ?? throw new ArgumentNullException(nameof(log));
            _fileSystem = fileSystem ?? throw new ArgumentNullException(nameof(fileSystem));
        }

        public override bool Execute()
        {
            try
            {
                _log.Debug("Running CreateExcelAddIn MSBuild Task");

                LogDiagnostics();

                RunSanityChecks();

                _dnaFilesToPack = new List<ITaskItem>();
                DnaFilesToPack = new ITaskItem[0];

                FilesInProject = FilesInProject ?? new ITaskItem[0];
                _log.Debug("Number of files in project: " + FilesInProject.Length);

                _configFilesInProject = GetConfigFilesInProject();
                _common = new BuildTaskCommon(FilesInProject, OutDirectory, FileSuffix32Bit, FileSuffix64Bit, ProjectName, AddInFileName);

                var buildItemsForDnaFiles = _common.GetBuildItemsForDnaFiles();

                TryCreateTlb();

                if (UnpackIsEnabled)
                    TryPublishDoc();

                TryBuildAddInFor32Bit(buildItemsForDnaFiles);

                _log.Information("---", MessageImportance.High);

                TryBuildAddInFor64Bit(buildItemsForDnaFiles);

                DnaFilesToPack = _dnaFilesToPack.ToArray();

                return true;
            }
            catch (Exception ex)
            {
                _log.Error(ex, ex.Message);
                _log.Error(ex, ex.ToString());
                return false;
            }
        }

        private void LogDiagnostics()
        {
            _log.Debug("----Arguments----");
            _log.Debug("FilesInProject: " + (FilesInProject ?? new ITaskItem[0]).Length);

            if (FilesInProject != null)
            {
                foreach (var f in FilesInProject)
                {
                    _log.Debug($"  {f.ItemSpec}");
                }
            }

            _log.Debug("OutDirectory: " + OutDirectory);
            _log.Debug("Xll32FilePath: " + Xll32FilePath);
            _log.Debug("Xll64FilePath: " + Xll64FilePath);
            _log.Debug("Create32BitAddIn: " + Create32BitAddIn);
            _log.Debug("Create64BitAddIn: " + Create64BitAddIn);
            _log.Debug("FileSuffix32Bit: " + FileSuffix32Bit);
            _log.Debug("FileSuffix64Bit: " + FileSuffix64Bit);
            _log.Debug("-----------------");
        }

        private void RunSanityChecks()
        {
            if (!_fileSystem.FileExists(Xll32FilePath))
            {
                throw new InvalidOperationException("File does not exist (Xll32FilePath): " + Xll32FilePath);
            }

            if (!_fileSystem.FileExists(Xll64FilePath))
            {
                throw new InvalidOperationException("File does not exist (Xll64FilePath): " + Xll64FilePath);
            }

            if (Create32BitAddIn && Create64BitAddIn && string.Equals(FileSuffix32Bit, FileSuffix64Bit, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("32-bit add-in suffix and 64-bit add-in suffix cannot be identical");
            }
        }

        private void TryCreateTlb()
        {
            if (!TlbCreate)
                return;

            string outputIntegrationDllPath = Path.Combine(OutDirectory, "ExcelDna.Integration.dll");
            bool outputIntegrationDllExists = File.Exists(outputIntegrationDllPath);
            if (!outputIntegrationDllExists)
                File.Copy(IntegrationDllPath, outputIntegrationDllPath);

            try
            {
                string outputFile = Path.Combine(OutDirectory, OutputFileName());
                string outputTlbFile = Path.ChangeExtension(outputFile, "tlb");
                if (TlbDscom)
                {
                    string args = $"tlbexport \"{outputFile}\"";
                    ProcessRunner.Run("dscom", args, "dscom", _log);

                    File.Delete(outputTlbFile);
                    File.Move(Path.GetFileName(outputTlbFile), outputTlbFile);
                }
                else
                {
                    string args = $"\"{outputFile}\" /out:\"{outputTlbFile}\"";
                    ProcessRunner.Run(TlbExp, args, "TlbExp", _log);
                }
            }
            finally
            {
                if (!outputIntegrationDllExists)
                    File.Delete(outputIntegrationDllPath);
            }
        }

        private void TryPublishDoc()
        {
            string docFilePath = GetDocPath();
            if (docFilePath == null)
                return;

            if (PackExcelAddIn.NoPublishPath(PublishPath))
                return;

            string destinationFolder = PackExcelAddIn.GetPublishDirectory(OutDirectory, PublishPath);
            Directory.CreateDirectory(destinationFolder);
            File.Copy(docFilePath, Path.Combine(destinationFolder, Path.GetFileName(docFilePath)), true);
        }

        private string GetDocPath()
        {
            string docFile = (AddInFileName ?? ProjectName + "-AddIn") + ".chm";
            string docFilePath = Path.Combine(OutDirectory, docFile);
            return File.Exists(docFilePath) ? docFilePath : null;
        }

        private ITaskItem[] GetConfigFilesInProject()
        {
            var configFilesInProject = FilesInProject
                .Where(file => string.Equals(Path.GetExtension(file.ItemSpec), ".config", StringComparison.OrdinalIgnoreCase))
                .OrderBy(file => file.ItemSpec)
                .ToArray();

            return configFilesInProject;
        }

        private void TryBuildAddInFor32Bit(BuildItemSpec[] buildItemsForDnaFiles)
        {
            foreach (var item in buildItemsForDnaFiles)
            {
                if (Create32BitAddIn && ShouldCopy32BitDnaOutput(item, buildItemsForDnaFiles))
                {
                    // Copy .dna file to build output folder for 32-bit
                    if (_fileSystem.FileExists(item.InputDnaFileName))
                        CopyFileToBuildOutput(item.InputDnaFileName, item.OutputDnaFileNameAs32Bit, overwrite: true);
                    else
                        WriteFileToBuildOutput(GetDefaultDnaText(), item.OutputDnaFileNameAs32Bit);

                    // Copy .xll file to build output folder for 32-bit
                    CopyFileToBuildOutput(Xll32FilePath, item.OutputXllFileNameAs32Bit, overwrite: true);

                    // Copy .config file to build output folder for 32-bit (if exist)
                    TryCopyConfigFileToOutput(item.InputConfigFileNameAs32Bit, item.InputConfigFileNameFallbackAs32Bit, item.OutputConfigFileNameAs32Bit);

                    if (!UnpackIsEnabled)
                        TryCopyDepsJsonToBuildOutput(item.OutputDnaFileNameAs32Bit);

                    AddDnaToListOfFilesToPack(item.OutputDnaFileNameAs32Bit, item.OutputXllFileNameAs32Bit, item.OutputConfigFileNameAs32Bit, Packed32BitXllName, "32");

                    if (UnpackIsEnabled)
                        PublishUnpackedAddin(item.OutputDnaFileNameAs32Bit, item.OutputXllFileNameAs32Bit);
                    else if (!CompressResources)
                        UncompressXll(item.OutputXllFileNameAs32Bit);
                }
            }
        }

        private void TryBuildAddInFor64Bit(BuildItemSpec[] buildItemsForDnaFiles)
        {
            foreach (var item in buildItemsForDnaFiles)
            {
                if (Create64BitAddIn && ShouldCopy64BitDnaOutput(item, buildItemsForDnaFiles))
                {
                    // Copy .dna file to build output folder for 64-bit
                    if (_fileSystem.FileExists(item.InputDnaFileName))
                        CopyFileToBuildOutput(item.InputDnaFileName, item.OutputDnaFileNameAs64Bit, overwrite: true);
                    else
                        WriteFileToBuildOutput(GetDefaultDnaText(), item.OutputDnaFileNameAs64Bit);

                    // Copy .xll file to build output folder for 64-bit
                    CopyFileToBuildOutput(Xll64FilePath, item.OutputXllFileNameAs64Bit, overwrite: true);

                    // Copy .config file to build output folder for 64-bit (if exist)
                    TryCopyConfigFileToOutput(item.InputConfigFileNameAs64Bit, item.InputConfigFileNameFallbackAs64Bit, item.OutputConfigFileNameAs64Bit);

                    if (!UnpackIsEnabled)
                        TryCopyDepsJsonToBuildOutput(item.OutputDnaFileNameAs64Bit);

                    AddDnaToListOfFilesToPack(item.OutputDnaFileNameAs64Bit, item.OutputXllFileNameAs64Bit, item.OutputConfigFileNameAs64Bit, Packed64BitXllName, "64");

                    if (UnpackIsEnabled)
                        PublishUnpackedAddin(item.OutputDnaFileNameAs64Bit, item.OutputXllFileNameAs64Bit);
                    else if (!CompressResources)
                        UncompressXll(item.OutputXllFileNameAs64Bit);
                }
            }
        }

        private static bool ShouldCopy32BitDnaOutput(BuildItemSpec item, IEnumerable<BuildItemSpec> buildItems)
        {
            if (item.InputDnaFileName.Equals(item.InputDnaFileNameAs32Bit))
            {
                return true;
            }

            var specificFileExists = buildItems
                .Any(bi => item.InputDnaFileNameAs32Bit.Equals(bi.InputDnaFileName, StringComparison.OrdinalIgnoreCase));

            return !specificFileExists;
        }

        private static bool ShouldCopy64BitDnaOutput(BuildItemSpec item, IEnumerable<BuildItemSpec> buildItems)
        {
            if (item.InputDnaFileName.Equals(item.InputDnaFileNameAs64Bit))
            {
                return true;
            }

            var specificFileExists = buildItems
                .Any(bi => item.InputDnaFileNameAs64Bit.Equals(bi.InputDnaFileName, StringComparison.OrdinalIgnoreCase));

            return !specificFileExists;
        }

        private void TryCopyConfigFileToOutput(string inputConfigFile, string inputFallbackConfigFile, string outputConfigFile)
        {
            var configFile = TryFindAppConfigFileName(inputConfigFile, inputFallbackConfigFile);
            if (!string.IsNullOrWhiteSpace(configFile))
            {
                CopyFileToBuildOutput(configFile, outputConfigFile, overwrite: true);
            }
        }

        private string TryFindAppConfigFileName(string preferredConfigFileName, string fallbackConfigFileName)
        {
            if (_configFilesInProject.Any(c => c.ItemSpec.Equals(preferredConfigFileName, StringComparison.OrdinalIgnoreCase)))
            {
                return preferredConfigFileName;
            }

            if (_configFilesInProject.Any(c => c.ItemSpec.Equals(fallbackConfigFileName, StringComparison.OrdinalIgnoreCase)))
            {
                return fallbackConfigFileName;

            }

            var appConfigFile = _configFilesInProject.FirstOrDefault(c => c.ItemSpec.Equals("App.config", StringComparison.OrdinalIgnoreCase));
            if (appConfigFile != null)
            {
                return appConfigFile.ItemSpec;
            }

            var linkedAppConfigFile = _configFilesInProject.FirstOrDefault(c => c.GetMetadata("Link").Equals("App.config", StringComparison.OrdinalIgnoreCase));
            if (linkedAppConfigFile != null)
            {
                return linkedAppConfigFile.ItemSpec;
            }

            return null;
        }

        private void TryCopyDepsJsonToBuildOutput(string outputDnaPath)
        {
            Integration.DnaLibrary dna = Integration.DnaLibrary.LoadFrom(File.ReadAllBytes(outputDnaPath), Path.GetDirectoryName(outputDnaPath));
            if (dna == null || dna.ExternalLibraries == null)
                return;

            foreach (Integration.ExternalLibrary ext in dna.ExternalLibraries)
            {
                string src = dna.ResolvePath(Path.ChangeExtension(ext.Path, "deps.json"));
                if (File.Exists(src))
                {
                    string dst = Path.ChangeExtension(outputDnaPath, "deps.json");
                    CopyFileToBuildOutput(src, dst, overwrite: true);
                    return;
                }
            }
        }

        private void CopyFileToBuildOutput(string sourceFile, string destinationFile, bool overwrite)
        {
            _log.Information(_fileSystem.GetRelativePath(sourceFile) + " -> " + _fileSystem.GetRelativePath(destinationFile));

            var destinationFolder = Path.GetDirectoryName(destinationFile);
            if (!string.IsNullOrWhiteSpace(destinationFolder) && !_fileSystem.DirectoryExists(destinationFolder))
            {
                _fileSystem.CreateDirectory(destinationFolder);
            }

            _fileSystem.CopyFile(sourceFile, destinationFile, overwrite);
        }

        private void WriteFileToBuildOutput(string sourceFileText, string destinationFile)
        {
            _log.Information(" -> " + _fileSystem.GetRelativePath(destinationFile));

            var destinationFolder = Path.GetDirectoryName(destinationFile);
            if (!string.IsNullOrWhiteSpace(destinationFolder) && !_fileSystem.DirectoryExists(destinationFolder))
            {
                _fileSystem.CreateDirectory(destinationFolder);
            }

            _fileSystem.WriteFile(sourceFileText, destinationFile);
        }

        private void AddDnaToListOfFilesToPack(string outputDnaFileName, string outputXllFileName, string outputXllConfigFileName, string packedFileName, string outputBitness)
        {
            if (!PackIsEnabled)
            {
                return;
            }

            string outputPackedXllFileName = PackExcelAddIn.GetOutputPackedXllFileName(outputXllFileName, packedFileName, PackedFileSuffix, PackExcelAddIn.GetPublishDirectory(OutDirectory, PublishPath));

            var metadata = new Hashtable
            {
                {"OutputDnaFileName", outputDnaFileName},
                {"OutputPackedXllFileName", outputPackedXllFileName},
                {"OutputXllConfigFileName", outputXllConfigFileName },
                {"OutputBitness", outputBitness },
                {"DocPath", GetDocPath() },
            };

            _dnaFilesToPack.Add(new TaskItem(outputDnaFileName, metadata));
        }

        private string GetDefaultDnaText()
        {
            string result = File.ReadAllText(TemplateDnaPath);

            {
                int startIndex = result.IndexOf("<!--");
                int endIndex = result.IndexOf("-->") + 3;
                result = result.Remove(startIndex, endIndex - startIndex);
            }

            if (!string.IsNullOrEmpty(AddInName))
                result = result.Replace("%ProjectName% Add-In", AddInName);
            else
                result = result.Replace("%ProjectName%", ProjectName);

            if (!string.IsNullOrWhiteSpace(AddInInclude))
            {
                string includes = "";
                foreach (string path in SplitDlls(AddInInclude))
                    includes += $"  <Reference Path=\"{path}\" Pack=\"true\" />" + Environment.NewLine;
                result = result.Replace("</DnaLibrary>", includes + "</DnaLibrary>");
            }

            if (DisableAssemblyContextUnload)
                result = result.Replace("<DnaLibrary ", "<DnaLibrary " + "DisableAssemblyContextUnload=\"true\" ");

            if (CustomRuntimeConfiguration)
                result = result.Replace("<DnaLibrary ", "<DnaLibrary " + $"CustomRuntimeConfiguration=\"{ProjectName}.runtimeconfig.json\" ");

            if (!string.IsNullOrWhiteSpace(RollForward))
            {
                result = result.Replace("<DnaLibrary ", "<DnaLibrary " + $"RollForward=\"{RollForward}\" ");
            }

            // For compatibility with .NET Framework 4 loader, we only set the exact version if its .NET 6+
            if (!TargetFrameworkVersion.StartsWith("v4."))
            {
                result = result.Replace(" RuntimeVersion=\"v4.0\"", $" RuntimeVersion=\"{TargetFrameworkVersion}\"");
            }

            result = UpdateExternalLibraries(result);

            return result;
        }

        private IEnumerable<string> SplitDlls(string dlls)
        {
            List<string> result = new List<string>();

            string outFiles = dlls.Replace(OutDirectory, "");
            foreach (string i in outFiles.Split(';'))
            {
                string path = i.Trim();
                if (path.Length > 0)
                    result.Add(path);
            }

            return result;
        }

        private string UpdateExternalLibraries(string dna)
        {
            int begin = dna.IndexOf("<ExternalLibrary ");
            int end = dna.IndexOf("/>", begin) + 2;
            string template = dna.Substring(begin, end - begin);

            string libraries = UpdateExternalLibrary(template, OutputFileName());
            if (!string.IsNullOrWhiteSpace(AddInExports))
            {
                foreach (string path in SplitDlls(AddInExports))
                    libraries += Environment.NewLine + $"  {UpdateExternalLibrary(template, path)}";
            }

            dna = dna.Remove(begin, end - begin);
            return dna.Insert(begin, libraries);
        }

        private string UpdateExternalLibrary(string template, string dllFileName)
        {
            string result = template;

            if (UseVersionAsOutputVersion)
                result = result.Replace("<ExternalLibrary ", "<ExternalLibrary " + "UseVersionAsOutputVersion=\"true\" ");

            if (ExplicitExports)
                result = result.Replace("ExplicitExports=\"false\"", "ExplicitExports=\"true\"");

            if (ExplicitRegistration)
                result = result.Replace("<ExternalLibrary ", "<ExternalLibrary " + "ExplicitRegistration=\"true\" ");

            if (ComServer)
                result = result.Replace("<ExternalLibrary ", "<ExternalLibrary " + "ComServer=\"true\" ");

            if (!LoadFromBytes)
                result = result.Replace("LoadFromBytes=\"true\"", "LoadFromBytes=\"false\"");

            return result.Replace("%OutputFileName%", dllFileName);
        }

        private string OutputFileName()
        {
            return !string.IsNullOrEmpty(AddInExternalLibraryPath) ? AddInExternalLibraryPath : TargetFileName;
        }

        private void PublishUnpackedAddin(string dnaPath, string xllPath)
        {
            var publishFolder = PackExcelAddIn.GetPublishDirectory(OutDirectory, PublishPath);
            Directory.CreateDirectory(publishFolder);
            UnpackXll(xllPath, new string[] { OutDirectory, publishFolder });

            if (PackExcelAddIn.NoPublishPath(PublishPath))
                return;

            List<string> filesToPublish = new List<string>();
            int result = PackedResources.ExcelDnaPack.Pack(dnaPath, null, false, false, false, null, filesToPublish, false, false, null, false, null, null, _log);
            if (result != 0)
                throw new ApplicationException($"Pack failed with exit code {result}.");
            foreach (string file in filesToPublish)
                File.Copy(file, Path.Combine(publishFolder, Path.GetFileName(file)), true);
        }

        private void UncompressXll(string xllPath)
        {
            string[] assemblies = { "ExcelDna.Integration", "ExcelDna.Loader" };
            foreach (var i in assemblies)
            {
                string name = i.ToUpperInvariant();
                UncompressResource(xllPath, name, ResourceHelper.TypeName.ASSEMBLY, "ASSEMBLY_LZMA");
                UncompressResource(xllPath, name, ResourceHelper.TypeName.PDB, "PDB_LZMA");
            }
        }

        private void UncompressResource(string xllPath, string name, ResourceHelper.TypeName typeName, string compressedTypeName)
        {
            byte[] data = LoadResource(xllPath, name, compressedTypeName);
            if (data == null)
                return;

            ResourceHelper.ResourceUpdater ru = new ResourceHelper.ResourceUpdater(Path.Combine(Directory.GetCurrentDirectory(), xllPath), false, _log);
            ru.AddFile(data, name, typeName, null, false, false);
            ru.RemoveResource(compressedTypeName, name);
            ru.EndUpdate();
        }

        private byte[] LoadResource(string xllPath, string name, string typeName)
        {
            IntPtr hModule = ResourceHelper.LoadXllResources(xllPath);
            if (hModule == IntPtr.Zero)
                throw new InvalidOperationException("Error loading resources from " + xllPath);

            try
            {
                return ResourceHelper.ResourceUpdater.LoadResourceBytes(hModule, typeName, name);
            }
            finally
            {
                ResourceHelper.FreeXllResources(hModule);
            }
        }

        private void UnpackXll(string xllPath, IEnumerable<string> destinationFolders)
        {
            string[] assemblies = { "ExcelDna.ManagedHost", "ExcelDna.Integration", "ExcelDna.Loader" };

            IntPtr hModule = ResourceHelper.LoadXllResources(xllPath);
            if (hModule == IntPtr.Zero)
                throw new InvalidOperationException("Error loading resources from " + xllPath);

            try
            {
                foreach (var i in assemblies)
                {
                    TryUnpackResource(hModule, i + ".dll", "ASSEMBLY", destinationFolders);
                    TryUnpackResource(hModule, i + ".pdb", "PDB", destinationFolders);
                }
            }
            finally
            {
                ResourceHelper.FreeXllResources(hModule);
            }

            foreach (var i in assemblies)
            {
                string name = i.ToUpperInvariant();
                TryRemoveResource(xllPath, name, "ASSEMBLY");
                TryRemoveResource(xllPath, name, "ASSEMBLY_LZMA");
                TryRemoveResource(xllPath, name, "PDB");
                TryRemoveResource(xllPath, name, "PDB_LZMA");
            }
        }

        private void TryUnpackResource(IntPtr hModule, string resourceFileName, string typeName, IEnumerable<string> destinationFolders)
        {
            byte[] data = ResourceHelper.ResourceUpdater.LoadResourceBytes(hModule, typeName, Path.GetFileNameWithoutExtension(resourceFileName).ToUpperInvariant());
            if (data != null)
            {
                foreach (string destinationFolder in destinationFolders)
                    File.WriteAllBytes(Path.Combine(destinationFolder, resourceFileName), data);
            }
        }

        private void TryRemoveResource(string xllPath, string name, string typeName)
        {
            var updater = new ResourceHelper.ResourceUpdater(xllPath, false, _log);
            try
            {
                updater.RemoveResource(typeName, name);
            }
            catch (System.ComponentModel.Win32Exception)
            {
                updater.EndUpdate(true);
                return;
            }
            updater.EndUpdate();
        }

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
        /// The file name of the primary output file for the build
        /// </summary>
        [Required]
        public string TargetFileName { get; set; }

        /// <summary>
        /// The version of the .NET that is required to run the add-in
        /// </summary>
        [Required]
        public string TargetFrameworkVersion { get; set; }

        /// <summary>
        /// The path to ExcelDna-Template.dna
        /// </summary>
        [Required]
        public string TemplateDnaPath { get; set; }

        /// <summary>
        /// The path to ExcelDna.Integration.dll
        /// </summary>
        [Required]
        public string IntegrationDllPath { get; set; }

        /// <summary>
        /// Compress (LZMA) of resources
        /// </summary>
        [Required]
        public bool CompressResources { get; set; }

        /// <summary>
        /// Controls how the add-in chooses a runtime when multiple runtime versions are available
        /// </summary>
        public string RollForward { get; set; }

        /// <summary>
        /// Enable/disable building 32-bit .dna files
        /// </summary>
        public bool Create32BitAddIn { get; set; }

        /// <summary>
        /// Enable/disable building 64-bit .dna files
        /// </summary>
        public bool Create64BitAddIn { get; set; }

        /// <summary>
        /// The name suffix for 32-bit .dna files
        /// </summary>
        public string FileSuffix32Bit { get; set; }

        /// <summary>
        /// The name suffix for 64-bit .dna files
        /// </summary>
        public string FileSuffix64Bit
        {
            get { return BuildTaskCommon.IsNone(_FileSuffix64Bit) ? null : _FileSuffix64Bit; }
            set { _FileSuffix64Bit = value; }
        }

        private string _FileSuffix64Bit;

        /// <summary>
        /// Enable/disable to have an .xll file with no packed assemblies
        /// </summary>
        public bool UnpackIsEnabled { get; set; }

        /// <summary>
        /// Enable/disable running ExcelDnaPack for .dna files
        /// </summary>
        public bool PackIsEnabled { get; set; }

        /// <summary>
        /// Packed add-in name suffix
        /// </summary>
        public string PackedFileSuffix { get; set; }

        /// <summary>
        /// Explicit 32-bit output file name
        /// </summary>
        public string Packed32BitXllName { get; set; }

        /// <summary>
        /// Explicit 64-bit output file name
        /// </summary>
        public string Packed64BitXllName { get; set; }

        /// <summary>
        /// Enable/disable cross-platform resource packing implementation when executing on Windows.
        /// </summary>
        public bool PackManagedOnWindows { get; set; }

        /// <summary>
        /// The output directory for the 'published' add-in
        /// </summary>
        public string PublishPath { get; set; }

        /// <summary>
        /// Custom add-in name
        /// </summary>
        public string AddInName { get; set; }

        /// <summary>
        /// Custom add-in file name
        /// </summary>
        public string AddInFileName { get; set; }

        /// <summary>
        /// Semicolon separated list of references written to the .dna file
        /// </summary>
        public string AddInInclude { get; set; }

        /// <summary>
        /// Semicolon separated external libraries to include in the .dna file
        /// </summary>
        public string AddInExports { get; set; }

        /// <summary>
        /// Custom path for ExternalLibrary
        /// </summary>
        public string AddInExternalLibraryPath { get; set; }

        /// <summary>
        /// Enable/disable collectible AssemblyLoadContext for .NET 6
        /// </summary>
        public bool DisableAssemblyContextUnload { get; set; }

        /// <summary>
        /// Enable/disable using the project's output runtimeconfig.json file for .NET 6
        /// </summary>
        public bool CustomRuntimeConfiguration { get; set; }

        /// <summary>
        /// Path to TlbExp.exe
        /// </summary>
        public string TlbExp { get; set; }

        /// <summary>
        /// Enable/disable .tlb file creation
        /// </summary>
        public bool TlbCreate { get; set; }

        /// <summary>
        /// Use Dscom instead of TlbExp
        /// </summary>
        [Required]
        public bool TlbDscom { get; set; }

        /// <summary>
        /// Replace XLL version information with data read from ExternalLibrary assembly
        /// </summary>
        public bool UseVersionAsOutputVersion { get; set; }

        /// <summary>
        /// Prevents every static public function from becomming a UDF, they will need an explicit [ExcelFunction] annotation
        /// </summary>
        public bool ExplicitExports { get; set; }

        /// <summary>
        /// Prevents automatic registration of functions and commands
        /// </summary>
        public bool ExplicitRegistration { get; set; }

        /// <summary>
        /// Enable/disable COM Server support
        /// </summary>
        public bool ComServer { get; set; }

        /// <summary>
        /// Enable/disable more dynamic .dll loading
        /// </summary>
        public bool LoadFromBytes { get; set; }

        /// <summary>
        /// The list of .dna files copied to the output
        /// </summary>
        [Output]
        public ITaskItem[] DnaFilesToPack { get; set; }
    }
}
