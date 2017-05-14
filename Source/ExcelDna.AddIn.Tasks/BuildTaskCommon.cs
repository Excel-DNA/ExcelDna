using System;
using System.IO;
using System.Linq;
using Microsoft.Build.Framework;

namespace ExcelDna.AddIn.Tasks
{
    public class BuildTaskCommon
    {
        private readonly string _outDirectory;
        private readonly string _fileSuffix32Bit;
        private readonly string _fileSuffix64Bit;
        private readonly ITaskItem[] _filesInProject;

        public BuildTaskCommon(ITaskItem[] filesInProject, string outDirectory, string fileSuffix32Bit, string fileSuffix64Bit)
        {
            if (filesInProject == null)
            {
                throw new ArgumentNullException("filesInProject");
            }

            _filesInProject = filesInProject;
            _outDirectory = outDirectory;
            _fileSuffix32Bit = fileSuffix32Bit;
            _fileSuffix64Bit = fileSuffix64Bit;
        }
        
        internal BuildItemSpec[] GetBuildItemsForDnaFiles()
        {
            var buildItemsForDnaFiles = (
                from item in _filesInProject
                where string.Equals(Path.GetExtension(item.ItemSpec), ".dna", StringComparison.OrdinalIgnoreCase)
                orderby item.ItemSpec
                let inputDnaFileNameAs32Bit = GetDnaFileNameAs32Bit(item.ItemSpec)
                let inputDnaFileNameAs64Bit = GetDnaFileNameAs64Bit(item.ItemSpec)
                select new BuildItemSpec
                {
                    InputDnaFileName = item.ItemSpec,

                    InputDnaFileNameAs32Bit = inputDnaFileNameAs32Bit,
                    InputDnaFileNameAs64Bit = inputDnaFileNameAs64Bit,

                    InputConfigFileNameAs32Bit = Path.ChangeExtension(inputDnaFileNameAs32Bit, ".config"),
                    InputConfigFileNameFallbackAs32Bit = GetAppConfigFileNameAs32Bit(),

                    InputConfigFileNameAs64Bit = Path.ChangeExtension(inputDnaFileNameAs64Bit, ".config"),
                    InputConfigFileNameFallbackAs64Bit = GetAppConfigFileNameAs64Bit(),

                    OutputDnaFileNameAs32Bit = Path.Combine(_outDirectory, inputDnaFileNameAs32Bit),
                    OutputDnaFileNameAs64Bit = Path.Combine(_outDirectory, inputDnaFileNameAs64Bit),

                    OutputXllFileNameAs32Bit = Path.Combine(_outDirectory, Path.ChangeExtension(inputDnaFileNameAs32Bit, ".xll")),
                    OutputXllFileNameAs64Bit = Path.Combine(_outDirectory, Path.ChangeExtension(inputDnaFileNameAs64Bit, ".xll")),

                    OutputConfigFileNameAs32Bit = Path.Combine(_outDirectory, Path.ChangeExtension(inputDnaFileNameAs32Bit, ".xll.config")),
                    OutputConfigFileNameAs64Bit = Path.Combine(_outDirectory, Path.ChangeExtension(inputDnaFileNameAs64Bit, ".xll.config")),
                }).ToArray();

            return buildItemsForDnaFiles;
        }

        private string GetDnaFileNameAs32Bit(string fileName)
        {
            return GetFileNameWithBitnessSuffix(fileName, _fileSuffix32Bit);
        }

        private string GetDnaFileNameAs64Bit(string fileName)
        {
            return GetFileNameWithBitnessSuffix(fileName, _fileSuffix64Bit);
        }

        private string GetAppConfigFileNameAs32Bit()
        {
            return GetFileNameWithBitnessSuffix("App.config", _fileSuffix32Bit);
        }

        private string GetAppConfigFileNameAs64Bit()
        {
            return GetFileNameWithBitnessSuffix("App.config", _fileSuffix64Bit);
        }

        private string GetFileNameWithBitnessSuffix(string fileName, string suffix)
        {
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName) ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(_fileSuffix32Bit))
            {
                var indexOfSuffix = fileNameWithoutExtension.LastIndexOf(_fileSuffix32Bit, StringComparison.OrdinalIgnoreCase);
                if (indexOfSuffix > 0)
                {
                    fileNameWithoutExtension = fileNameWithoutExtension.Remove(indexOfSuffix);
                }
            }

            if (!string.IsNullOrWhiteSpace(_fileSuffix64Bit))
            {
                var indexOfSuffix = fileNameWithoutExtension.LastIndexOf(_fileSuffix64Bit, StringComparison.OrdinalIgnoreCase);
                if (indexOfSuffix > 0)
                {
                    fileNameWithoutExtension = fileNameWithoutExtension.Remove(indexOfSuffix);
                }
            }

            var extension = Path.GetExtension(fileName);

            return Path.Combine(Path.GetDirectoryName(fileName) ?? string.Empty, fileNameWithoutExtension + suffix + extension);
        }
    }
}
