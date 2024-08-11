namespace ExcelDna.AddIn.RuntimeTests
{
    internal static class Logger
    {
        static string log = "";
        static readonly object logLock = new object();

        public static void Log(string message)
        {
            lock (logLock)
            {
                log += message + Environment.NewLine;
            }
        }

        public static string GetLog()
        {
            lock (logLock)
            {
                return log;
            }
        }

        public static void ClearLog()
        {
            lock (logLock)
            {
                log = "";
            }
        }
    }
}
