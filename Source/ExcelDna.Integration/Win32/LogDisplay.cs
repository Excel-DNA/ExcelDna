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

        public static void RecordLine(string s)
        {
            log += s + Environment.NewLine;

            if (window != null)
                window.SetText(log);
        }

        private static LogDisplayWindow window;
        private static string log = "";
    }
}

#endif
