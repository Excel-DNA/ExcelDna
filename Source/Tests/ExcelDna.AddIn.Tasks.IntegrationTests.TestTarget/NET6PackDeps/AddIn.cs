using System;
using System.Data.SQLite;
using System.IO;
using System.Management;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace NET6PackDeps
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            var thisAddInName = Path.GetFileName((string)XlCall.Excel(XlCall.xlGetName));
            try
            {
                var message = string.Format("Excel-DNA Add-In '{0}' loaded!", thisAddInName);
                message += Environment.NewLine + SQLiteEvaluate();
                message += Environment.NewLine + GetComputerName();
                MessageBox.Show(message, thisAddInName, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), thisAddInName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void AutoClose()
        {
        }

        private static string SQLiteEvaluate()
        {
            SQLiteConnection _connection = new SQLiteConnection("DataSource=:memory:");
            _connection.Open();

            SQLiteCommand cmd = _connection.CreateCommand();
            cmd.CommandText = "SELECT 40 + 2";
            var result = cmd.ExecuteScalar();
            return result.ToString();
        }

        public static string GetComputerName()
        {
            ManagementObjectSearcher searcher =
                new ManagementObjectSearcher("root\\CIMV2",
                "SELECT Name FROM Win32_ComputerSystem");

            foreach (ManagementObject queryObj in searcher.Get())
            {
                return queryObj["Name"].ToString();
            }

            return "No Name !?";
        }
    }
}
