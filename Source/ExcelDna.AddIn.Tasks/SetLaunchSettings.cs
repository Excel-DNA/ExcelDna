using System;
using System.IO;
using System.Linq;
using ExcelDna.AddIn.Tasks.Logging;
using Microsoft.Build.Framework;
using ExcelDna.AddIn.Tasks.Utils;

namespace ExcelDna.AddIn.Tasks
{
    public class SetLaunchSettings : ExcelDebugTask
    {
        public SetLaunchSettings()
        {
            _log = new BuildLogger(this, "ExcelDnaSetLaunchSettings");
            _excelDetectorLazy = new Lazy<IExcelDetector>(() => new ExcelDetector(_log));
        }

        public override bool Execute()
        {
            try
            {
                _log.Debug("Running SetLaunchSettings MSBuild Task");

                string settingsPath = Path.Combine(ProjectDirectory, "Properties", "launchSettings.json");
                if (File.Exists(settingsPath))
                    return true;

                _excelDetector = _excelDetectorLazy.Value;
                var excelExePath = GetExcelPath();
                var addInForDebugging = GetAddInForDebugging(excelExePath);

                var settings = new
                {
                    profiles = new
                    {
                        ExcelDna = new
                        {
                            commandName = "Executable",
                            executablePath = excelExePath,
                            commandLineArgs = Path.GetFileName(addInForDebugging)
                        }
                    }
                };

                Directory.CreateDirectory(Path.GetDirectoryName(settingsPath));
                File.WriteAllText(settingsPath, Newtonsoft.Json.JsonConvert.SerializeObject(settings, Newtonsoft.Json.Formatting.Indented));
            }
            catch (Exception ex)
            {
                _log.Warning(ex, ex.Message);
            }

            return true;
        }

        /// <summary>
        /// The absolute path of the directory where the project file is located
        /// </summary>
        [Required]
        public string ProjectDirectory { get; set; }
    }
}
