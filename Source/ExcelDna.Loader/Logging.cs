using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace ExcelDna.Loader.Logging
{
    // The NLog levels
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
        string _logPath;
        public FileLogTarget(string pathXll)
        {
            // TODO: Do we just append? Forever?
            _logPath = pathXll + ".log";
        }

        public void Write(string message)
        {
            File.AppendAllText(_logPath, message);
        }
    }

    // For now a static class is simple and convenient - we only do Registration Logging
    // in future we might split this into an ILogger, formatting etc....
    // Or we implement a simple interface like CM, and allow user plug-ins
    // Or just implement it via the System.Diagnostics.Trace...
    internal static class RegistrationLogger
    {
        // Name of the logger
        // const string _name = "Registration";
        static LogLevel _level = LogLevel.Off;
        static FileLogTarget _target = null;
        static string _indentPrefix = "";

        // public static TraceSource TraceSource = new TraceSource("Registration");

        public static void Initialize(string pathXll)
        {
            try
            {
                string logConfig = System.Configuration.ConfigurationManager.AppSettings["RegistrationLogLevel"];
                if (logConfig != null)
                {
                    _level = (LogLevel)Enum.Parse(typeof(LogLevel), logConfig);
                    // set _logLevel, _target etc.
                    _target = new FileLogTarget(pathXll);
                    Log(LogLevel.Trace, "RegistrationLogger.Initialize");
                }
            }
            catch (Exception)
            {
                // Suppress any trouble here.
                _level = LogLevel.Off;
                _target = null;
            }
        }

        public static void Log(LogLevel level, string message, params object[] args)
        {
            Debug.Write(string.Format("{0:yyyy-MM-dd HH:mm:ss} {1} {2}\r\n", DateTime.Now, level, string.Format(message, args)));

            if (level >= _level)
            {
                _target.Write(string.Format("{0:yyyy-MM-dd HH:mm:ss} {1} {2}\r\n", DateTime.Now, level, string.Format(message, args)));
            }
        }

        public static void Info(string message, params object[] args)
        {
            Log(LogLevel.Info, message, args);
        }

        public static void Warn(string message, params object[] args)
        {
            Log(LogLevel.Warn, message, args);
        }

        public static void Error(string message, params object[] args)
        {
            Log(LogLevel.Error, message, args);
        }

        public static void ErrorException(string message, Exception ex)
        {
            Log(LogLevel.Error, "{0} : {1} - {2}", message, ex.GetType().Name.ToString(), ex.Message);
        }

        public static void Indent()
        {
            _indentPrefix += "\t";
        }

        public static void Unindent()
        {
            if (_indentPrefix.Length > 0)
            {
                _indentPrefix = _indentPrefix.Substring(0, _indentPrefix.Length - 1);
            }
        }
    }

    static class LogManager
    {
        public static void Initialize(string xllPath)
        {
            RegistrationLogger.Initialize(xllPath);
            RegistrationLogger.Log(LogLevel.Trace, "Hello {0}", "World!");
        }
    }
}
