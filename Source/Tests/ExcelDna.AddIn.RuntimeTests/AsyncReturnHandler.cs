using ExcelDna.Integration;
using ExcelDna.Registration;

namespace ExcelDna.AddIn.RuntimeTests
{
    internal class AsyncReturnHandler : FunctionExecutionHandler
    {
        public override void OnSuccess(FunctionExecutionArgs args)
        {
            if (args.FunctionName == "MyAsyncGettingData" && args.ReturnValue.Equals(ExcelError.ExcelErrorNA))
                args.ReturnValue = ExcelError.ExcelErrorGettingData;
        }

        [ExcelFunctionExecutionHandlerSelector]
        public static IFunctionExecutionHandler AsyncReturnHandlerSelector(IExcelFunctionInfo functionInfo)
        {
            return new AsyncReturnHandler();
        }
    }
}
