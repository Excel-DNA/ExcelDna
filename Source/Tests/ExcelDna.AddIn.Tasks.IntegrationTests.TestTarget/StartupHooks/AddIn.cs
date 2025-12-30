using ExcelDna.Integration;

namespace StartupHooks
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            var thisAddInName = Path.GetFileName((string)XlCall.Excel(XlCall.xlGetName));
            bool hooksLoaded = AppDomain.CurrentDomain.GetAssemblies().Any(i => i.GetName().Name == "dotnet-watch");
            var message = string.Format("Excel-DNA Add-In '{0}' loaded!", thisAddInName);
            message += Environment.NewLine + Environment.NewLine + $"Hooks loaded: {hooksLoaded}.";

            MessageBox.Show(message, thisAddInName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void AutoClose()
        {
        }
    }
}
