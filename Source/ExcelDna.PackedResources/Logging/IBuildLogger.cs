using System;

namespace ExcelDna.PackedResources.Logging
{
    public interface IBuildLogger
    {
        void Message(LogImportance importance, string format, params object[] args);
        void Verbose(string format, params object[] args);
        void Debug(string format, params object[] args);
        void Information(string format, params object[] args);

        void Warning(Exception exception, string format, params object[] args);
        void Warning(string code, string format, params object[] args);

        void Error(Exception exception, string format, params object[] args);
        void Error(Type errorSource, string format, params object[] args);
        void Error(string code, string format, params object[] args);
    }
}
