using ExcelDna.Integration;

namespace ExcelDna.AddIn.RuntimeTests
{
    internal class FunctionLoggingHandler : FunctionExecutionHandler
    {
        public override void OnEntry(FunctionExecutionArgs args)
        {
            // FunctionExecutionArgs gives access to the function name and parameters,
            // and gives some options for flow redirection.

            // Tag will flow through the whole handler
            args.Tag = args.FunctionName;
            Logger.Log($"{args.Tag} - OnEntry - Args: {args.Arguments.Select(arg => arg.ToString())}");
        }

        public override void OnSuccess(FunctionExecutionArgs args)
        {
            Logger.Log($"{args.Tag} - OnSuccess - Result: {args.ReturnValue}");
        }

        public override void OnException(FunctionExecutionArgs args)
        {
            Logger.Log($"{args.Tag} - OnException - Message: {args.Exception}");
        }

        public override void OnExit(FunctionExecutionArgs args)
        {
            Logger.Log($"{args.Tag} - OnExit");
        }

        [ExcelFunctionExecutionHandlerSelector]
        public static IFunctionExecutionHandler LoggingHandlerSelector(IExcelFunctionInfo functionInfo)
        {
            return new FunctionLoggingHandler();
        }
    }
}
