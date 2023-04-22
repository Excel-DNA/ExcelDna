using System.IO;
using System.Reflection;
using System.Runtime.Loader;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace NET6DisableAssemblyContextUnload
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            bool collectible = AssemblyLoadContext.GetLoadContext(Assembly.GetExecutingAssembly())!.IsCollectible;
            var thisAddInName = Path.GetFileName((string)XlCall.Excel(XlCall.xlGetName));
            var message = $"Excel-DNA Add-In '{thisAddInName}' loaded! IsCollectible={collectible}.";

            MessageBox.Show(message, thisAddInName, MessageBoxButtons.OK, collectible ? MessageBoxIcon.Error : MessageBoxIcon.Information);
        }

        public void AutoClose()
        {
        }
    }
}
