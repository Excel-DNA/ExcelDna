using Microsoft.Build.Framework;

namespace ExcelDna.AddIn.Tasks
{
    public abstract class AbstractTask : ITask
    {
        public abstract bool Execute();

        protected void LogMessage(string message, MessageImportance importance = MessageImportance.High)
        {
            BuildEngine.LogMessageEvent(new BuildMessageEventArgs("ExcelDnaBuild: " + message, "ExcelDnaBuild", "ExcelDnaBuild", importance));
        }

        protected void LogWarning(string code, string message)
        {
            BuildEngine.LogWarningEvent(new BuildWarningEventArgs("ExcelDnaBuild", code, null, 0, 0, 0, 0, message, "ExcelDnaBuild", "ExcelDnaBuild"));
        }

        protected void LogError(string code, string message)
        {
            BuildEngine.LogErrorEvent(new BuildErrorEventArgs("ExcelDnaBuild", code, null, 0, 0, 0, 0, message, "ExcelDnaBuild", "ExcelDnaBuild"));
        }

        public IBuildEngine BuildEngine { get; set; }
        public ITaskHost HostObject { get; set; }
    }
}
