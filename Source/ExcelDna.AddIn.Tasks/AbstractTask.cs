using System;
using Microsoft.Build.Framework;

namespace ExcelDna.AddIn.Tasks
{
    public abstract class AbstractTask : ITask
    {
        private readonly string _targetName;

        protected AbstractTask(string targetName)
        {
            if (string.IsNullOrWhiteSpace(targetName))
            {
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(targetName));
            }

            _targetName = targetName;
        }

        public abstract bool Execute();

        protected void LogDebugMessage(string message)
        {
            #if DEBUG
                LogMessage(message, MessageImportance.High);
            #else
                LogMessage(message, MessageImportance.Low);
            #endif
        }

        protected void LogMessage(string message, MessageImportance importance = MessageImportance.High)
        {
            BuildEngine.LogMessageEvent(new BuildMessageEventArgs(string.Format("{0}: {1}", _targetName, message),
                _targetName, _targetName, importance));
        }

        protected void LogWarning(string code, string message)
        {
            BuildEngine.LogWarningEvent(new BuildWarningEventArgs(_targetName, code, null, 0, 0, 0, 0, message,
                _targetName, _targetName));
        }

        protected void LogError(string code, string message)
        {
            BuildEngine.LogErrorEvent(new BuildErrorEventArgs(_targetName, code, null, 0, 0, 0, 0, message, _targetName,
                _targetName));
        }

        public IBuildEngine BuildEngine { get; set; }
        public ITaskHost HostObject { get; set; }
    }
}
