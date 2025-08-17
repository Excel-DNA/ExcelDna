using ExcelDna.Integration;

namespace ExcelDna.AddIn.RuntimeTests
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            DynamicFunctions.Register();
        }

        public void AutoClose()
        {
        }
    }
}
