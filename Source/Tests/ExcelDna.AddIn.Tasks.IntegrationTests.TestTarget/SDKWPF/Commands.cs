using ExcelDna.Integration;
using System;
using System.Windows;

namespace SDKWPF
{
    public class Commands
    {
        [ExcelCommand(MenuText = "OpenWindow")]
        public static void OpenWindow()
        {
            try
            {
                ShowWindow();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private static void ShowWindow()
        {
            Window1 window1 = new Window1();
            window1.ShowDialog();
        }
    }
}
