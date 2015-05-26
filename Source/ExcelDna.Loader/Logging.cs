using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelDna.Loader.Logging
{
    // Matches NLog
    internal enum LogLevel
    {
        Trace = 0,
        Debug = 1,
        Info = 2,
        Warn = 3,
        Error = 4,
        Fatal = 5,
        Off = 6
    }

    // A log target that writes to a file
    internal class FileLogTarget
    {
        string _fileName;
        public FileLogTarget(string pathXll)
        {
            // TODO: Do we just append? Forever?
            _fileName = pathXll + ".log";
        }

        public void Write(string message)
        {
            File.AppendText(message);
        }
    }

    // For now a static class is simple and convenient - we only do Registration Logging
    // in future we might split this into an ILogger, formatting etc....
    // Or we implement a simple interface like CM, and allow user plug-ins
    internal static class RegistrationLogger
    {
        // Name of the logger
        const string _name = "Registration";
        static LogLevel _level = LogLevel.Off;
        static FileLogTarget _target;

        public static void Initialize(string pathXll)
        {
            try
            {
                // TODO: Load configuration from .config file
                // set _logLevel, _target etc.
                _target = new FileLogTarget(pathXll);
            }
            catch (Exception)
            {
                // Suppress any trouble here.
                _target = null;
                _level = LogLevel.Off;
            }
        }

        public static void Log(LogLevel level, string message, params object[] args)
        {
            if (level >= _level)
            {
                _target.Write(string.Format("{0:yyyy-MM-dd HH:mm:ss} {1} {2}\r\n", DateTime.Now, level, string.Format(message, args)));
            }
        }
    }

    static class LogManager
    {
        public static void Initialize(string xllPath)
        {
            RegistrationLogger.Initialize(xllPath);
        }
    }
}
