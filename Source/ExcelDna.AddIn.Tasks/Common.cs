using System;
using System.IO;
using System.Linq;
using Microsoft.Build.Framework;

namespace ExcelDna.AddIn.Tasks
{
    public class BuildTaskCommon
    {
        private string OutDirectory;
        private string FileSuffix32Bit;
        private string FileSuffix64Bit;
        private ITaskItem[] FilesInProject;

        public BuildTaskCommon(ITaskItem[] FilesInProject, string OutDirectory, string FileSuffix32Bit, string FileSuffix64Bit)
        {
            this.FilesInProject = FilesInProject;
            this.OutDirectory = OutDirectory;
            this.FileSuffix32Bit = FileSuffix32Bit;
            this.FileSuffix64Bit = FileSuffix64Bit;
        }
        
        internal BuildItemSpec[] GetBuildItemsForDnaFiles()
        {
            var buildItemsForDnaFiles = (
                from item in FilesInProject
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

                    OutputDnaFileNameAs32Bit = Path.Combine(OutDirectory, inputDnaFileNameAs32Bit),
                    OutputDnaFileNameAs64Bit = Path.Combine(OutDirectory, inputDnaFileNameAs64Bit),

                    OutputXllFileNameAs32Bit = Path.Combine(OutDirectory, Path.ChangeExtension(inputDnaFileNameAs32Bit, ".xll")),
                    OutputXllFileNameAs64Bit = Path.Combine(OutDirectory, Path.ChangeExtension(inputDnaFileNameAs64Bit, ".xll")),

                    OutputConfigFileNameAs32Bit = Path.Combine(OutDirectory, Path.ChangeExtension(inputDnaFileNameAs32Bit, ".xll.config")),
                    OutputConfigFileNameAs64Bit = Path.Combine(OutDirectory, Path.ChangeExtension(inputDnaFileNameAs64Bit, ".xll.config")),
                }).ToArray();

            return buildItemsForDnaFiles;
        }

        private string GetDnaFileNameAs32Bit(string fileName)
        {
            return GetFileNameWithBitnessSuffix(fileName, FileSuffix32Bit);
        }

        private string GetDnaFileNameAs64Bit(string fileName)
        {
            return GetFileNameWithBitnessSuffix(fileName, FileSuffix64Bit);
        }

        private string GetAppConfigFileNameAs32Bit()
        {
            return GetFileNameWithBitnessSuffix("App.config", FileSuffix32Bit);
        }

        private string GetAppConfigFileNameAs64Bit()
        {
            return GetFileNameWithBitnessSuffix("App.config", FileSuffix64Bit);
        }

        private string GetFileNameWithBitnessSuffix(string fileName, string suffix)
        {
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName) ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(FileSuffix32Bit))
            {
                var indexOfSuffix = fileNameWithoutExtension.LastIndexOf(FileSuffix32Bit, StringComparison.OrdinalIgnoreCase);
                if (indexOfSuffix > 0)
                {
                    fileNameWithoutExtension = fileNameWithoutExtension.Remove(indexOfSuffix);
                }
            }

            if (!string.IsNullOrWhiteSpace(FileSuffix64Bit))
            {
                var indexOfSuffix = fileNameWithoutExtension.LastIndexOf(FileSuffix64Bit, StringComparison.OrdinalIgnoreCase);
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
