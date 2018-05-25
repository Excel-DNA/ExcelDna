using System;
using Microsoft.Build.Framework;

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

        public void Message(MessageImportance importance, string format, params object[] args)
        {
            _buildTask.BuildEngine.LogMessageEvent(new BuildMessageEventArgs($"{_targetName}: {string.Format(format, args)}",
                _targetName, _targetName, importance));
        }

        public void Verbose(string format, params object[] args)
        {
            Message(MessageImportance.Low, format);
        }

        public void Debug(string format, params object[] args)
        {
            Message(MessageImportance.Normal, format);
        }

        public void Information(string format, params object[] args)
        {
            Message(MessageImportance.High, format);
        }

        public void Warning(Exception exception, string format, params object[] args)
        {
            if (exception == null) throw new ArgumentNullException(nameof(exception));

            Warning("DNA" + exception.GetType().Name.GetHashCode(), format, args);
        }

        public void Warning(string code, string format, params object[] args)
        {
            _buildTask.BuildEngine.LogWarningEvent(new BuildWarningEventArgs(_targetName, code, null, 0, 0, 0, 0,
                string.Format(format, args), _targetName, _targetName));
        }

        public void Error(Exception exception, string format, params object[] args)
        {
            if (exception == null) throw new ArgumentNullException(nameof(exception));

            Error("DNA" + exception.GetType().Name.GetHashCode(), format, args);
        }

        public void Error(string code, string format, params object[] args)
        {
            _buildTask.BuildEngine.LogErrorEvent(new BuildErrorEventArgs(_targetName, code, null, 0, 0, 0, 0,
                string.Format(format, args), _targetName, _targetName));
        }
    }
}
