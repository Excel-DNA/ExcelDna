using System;
using Microsoft.Build.Framework;

namespace ExcelDna.AddIn.Tasks.Logging
{
    public interface IBuildLogger
    {
        void Message(MessageImportance importance, string format, params object[] args);
        void Verbose(string format, params object[] args);
        void Debug(string format, params object[] args);
        void Information(string format, params object[] args);

        void Warning(Exception exception, string format, params object[] args);
        void Warning(string code, string format, params object[] args);

        void Error(Exception exception, string format, params object[] args);
        void Error(string code, string format, params object[] args);
    }
}
