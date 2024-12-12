using System;
using Microsoft.Build.Framework;
using ExcelDna.PackedResources.Logging;

namespace ExcelDna.AddIn.Tasks.Logging
{
    internal class BuildLogger : IBuildLogger
    {
        private readonly ITask _buildTask;
        private readonly string _targetName;

        public BuildLogger(ITask buildTask, string targetName)
        {
            if (string.IsNullOrWhiteSpace(targetName))
            {
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(targetName));
            }

            _buildTask = buildTask ?? throw new ArgumentNullException(nameof(buildTask));
            _targetName = targetName;
        }

        public void Message(LogImportance importance, string format, params object[] args)
        {
            _buildTask.BuildEngine.LogMessageEvent(new BuildMessageEventArgs($"{_targetName}: {string.Format(format, args)}",
                _targetName, _targetName, (MessageImportance)importance));
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
            _buildTask.BuildEngine.LogWarningEvent(new BuildWarningEventArgs(_targetName, code, null, 0, 0, 0, 0,
                string.Format(format, args), _targetName, _targetName));
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
            _buildTask.BuildEngine.LogErrorEvent(new BuildErrorEventArgs(_targetName, code, null, 0, 0, 0, 0,
                string.Format(format, args), _targetName, _targetName));
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
