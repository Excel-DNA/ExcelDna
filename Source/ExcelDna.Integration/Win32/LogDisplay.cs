#if !USE_WINDOWS_FORMS

using System;

namespace ExcelDna.Integration.Win32
{
    internal class LogDisplay
    {
        public static void Show()
        {
            if (window == null)
                window = new LogDisplayWindow(new CreateParams(), log);

            window.Show();
        }

        public static void RecordLine(string s, params object[] args)
        {
            string message = args.Length > 0 ? string.Format(s, args) : s;
            log += message + Environment.NewLine;

            if (window != null)
                window.SetText(log);
        }

        public static void WriteLine(string format, params object[] args)
        {
            Show();
            RecordLine(format, args);
        }

        private static LogDisplayWindow window;
        private static string log = "";
    }
}

#endif
