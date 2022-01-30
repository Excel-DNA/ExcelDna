using System;
using System.IO;
using System.Linq;
using ExcelDna.AddIn.Tasks.Logging;
using Microsoft.Build.Framework;
using ExcelDna.AddIn.Tasks.Utils;
using Newtonsoft.Json.Linq;

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

                _excelDetector = _excelDetectorLazy.Value;
                var excelExePath = GetExcelPath();
                var addInForDebugging = GetAddInForDebugging(excelExePath);
                var profile = new
                {
                    commandName = "Executable",
                    executablePath = excelExePath,
                    commandLineArgs = Path.GetFileName(addInForDebugging)
                };

                if (File.Exists(settingsPath))
                    Update(settingsPath, profile);
                else
                    Create(settingsPath, profile);
            }
            catch (Exception ex)
            {
                _log.Warning(ex, ex.Message);
            }

            return true;
        }

        private void Create(string settingsPath, object profile)
        {
            var settings = new
            {
                profiles = new
                {
                    ExcelDna = profile
                }
            };

            Directory.CreateDirectory(Path.GetDirectoryName(settingsPath));
            File.WriteAllText(settingsPath, Newtonsoft.Json.JsonConvert.SerializeObject(settings, Newtonsoft.Json.Formatting.Indented));
        }

        private void Update(string settingsPath, object profile)
        {
            JObject settings = JObject.Parse(File.ReadAllText(settingsPath));
            JObject profiles = (JObject)settings["profiles"];
            profiles["ExcelDna"] = JToken.FromObject(profile);
            File.WriteAllText(settingsPath, settings.ToString());
        }

        /// <summary>
        /// The absolute path of the directory where the project file is located
        /// </summary>
        [Required]
        public string ProjectDirectory { get; set; }
    }
}
