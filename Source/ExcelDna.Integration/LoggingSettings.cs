using Microsoft.Win32;
using System;
using System.Diagnostics;

namespace ExcelDna.Logging
{
    internal class LoggingSettings
    {
        public SourceLevels SourceLevel { get; }
        public TraceEventType? LogDisplayLevel { get; }
        public TraceEventType? DebuggerLevel { get; }

        public LoggingSettings()
        {
            SourceLevel = Enum.TryParse(GetCustomSetting("SOURCE_LEVEL", "SourceLevel"), out SourceLevels sourceLevelResult) ? sourceLevelResult : SourceLevels.Warning;

            if (Enum.TryParse(GetCustomSetting("LOGDISPLAY_LEVEL", "LogDisplayLevel"), out TraceEventType logDisplayLevelResult))
                LogDisplayLevel = logDisplayLevelResult;

            if (Enum.TryParse(GetCustomSetting("DEBUGGER_LEVEL", "DebuggerLevel"), out TraceEventType debuggerLevelResult))
                DebuggerLevel = debuggerLevelResult;
        }

        private static string GetCustomSetting(string environmentName, string registryName)
        {
            return Environment.GetEnvironmentVariable($"EXCELDNA_DIAGNOSTICS_{environmentName}") ??
                (Registry.GetValue(@"HKEY_CURRENT_USER\Software\ExcelDna\Diagnostics", registryName, null) as string ??
                Registry.GetValue(@"HKEY_LOCAL_MACHINE\Software\ExcelDna\Diagnostics", registryName, null) as string);
        }
    }
}
