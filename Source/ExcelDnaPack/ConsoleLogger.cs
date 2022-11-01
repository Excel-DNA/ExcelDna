using System;
using ExcelDna.PackedResources.Logging;

namespace ExcelDna.AddIn.Tasks.Logging
{
    internal class ConsoleLogger : IBuildLogger
    {
        private readonly string _targetName;

        public ConsoleLogger(string targetName)
        {
            if (string.IsNullOrWhiteSpace(targetName))
            {
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(targetName));
            }

            _targetName = targetName;
        }

        public void Message(LogImportance importance, string format, params object[] args)
        {
            Console.WriteLine($"{_targetName}: {string.Format(format, args)}",
                _targetName, _targetName, importance);
        }

        public void Verbose(string format, params object[] args)
        {
            Message(LogImportance.Low, format, args);
        }

        public void Debug(string format, params object[] args)
        {
            Message(LogImportance.Normal, format, args);
        }

        public void Information(string format, params object[] args)
        {
            Message(LogImportance.High, format, args);
        }

        public void Warning(Exception exception, string format, params object[] args)
        {
            if (exception == null) throw new ArgumentNullException(nameof(exception));

            Warning(GetErrorCode(exception), format, args);
        }

        public void Warning(string code, string format, params object[] args)
        {
            Message(LogImportance.Normal, $"{code}:{format}", args);
        }

        public void Error(Exception exception, string format, params object[] args)
        {
            if (exception == null) throw new ArgumentNullException(nameof(exception));

            Error(GetErrorCode(exception), format, args);
        }

        public void Error(Type errorSource, string format, params object[] args)
        {
            Error(GetErrorCode(errorSource), format, args);
        }

        public void Error(string code, string format, params object[] args)
        {
            Message(LogImportance.High, $"{code}:{format}", args);
        }

        private static string GetErrorCode(Exception exception)
        {
            return GetErrorCode(exception.GetType());
        }

        private static string GetErrorCode(Type type)
        {
            return "DNA" + type.Name.GetHashCode();
        }
    }
}
