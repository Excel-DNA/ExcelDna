using System;
using System.IO;
using System.Threading.Channels;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace NET6Channels
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            var thisAddInName = Path.GetFileName((string)XlCall.Excel(XlCall.xlGetName));
            var message = string.Format("Excel-DNA Add-In '{0}' loaded!", thisAddInName);

            var channel = Channel.CreateUnbounded<int>();

            message += Environment.NewLine;
            message += Environment.NewLine;
            message += $"channel.Reader.Count {channel.Reader.Count}.";
            message += Environment.NewLine;
            message += $"{typeof(Channel).Assembly.FullName}";

            MessageBox.Show(message, thisAddInName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void AutoClose()
        {
        }
    }
}
