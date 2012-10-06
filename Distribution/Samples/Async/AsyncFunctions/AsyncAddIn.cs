using ExcelDna.Integration;

namespace AsyncFunctions
{
    public class AsyncTestAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelAsyncUtil.Initialize();
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex => "!!! EXCEPTION: " + ex.ToString());
        }

        public void AutoClose()
        {
        }
    }

}
