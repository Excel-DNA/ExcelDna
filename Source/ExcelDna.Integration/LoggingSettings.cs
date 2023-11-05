using Microsoft.Win32;
using System;
using System.Diagnostics;

namespace ExcelDna.Logging
{
    internal class LoggingSettings
    {
        public SourceLevels SourceLevel { get; }

        public LoggingSettings()
        {
            SourceLevel = Enum.TryParse(GetCustomSetting("SOURCE_LEVEL", "SourceLevel"), out SourceLevels result) ? result : SourceLevels.Warning;
        }

        private static string GetCustomSetting(string environmentName, string registryName)
        {
            return Environment.GetEnvironmentVariable($"EXCELDNA_DIAGNOSTICS_{environmentName}") ??
                (Registry.GetValue(@"HKEY_CURRENT_USER\Software\ExcelDna\Diagnostics", registryName, null) as string ??
                Registry.GetValue(@"HKEY_LOCAL_MACHINE\Software\ExcelDna\Diagnostics", registryName, null) as string);
        }
    }
}
