using ExcelDna.Integration;
using ExcelDna.Registration;

namespace ExcelDna.AddIn.RuntimeTests
{
    [AttributeUsage(AttributeTargets.Method)]
    public class LoggingAttribute : Attribute
    {
        public int ID { get; set; }

        public LoggingAttribute(int ID)
        {
            this.ID = ID;
        }
    }

    internal class FunctionLoggingHandler : FunctionExecutionHandler
    {
        public int? ID { get; set; }

        public override void OnEntry(FunctionExecutionArgs args)
        {
            // FunctionExecutionArgs gives access to the function name and parameters,
            // and gives some options for flow redirection.

            // Tag will flow through the whole handler
            if (ID.HasValue)
                args.Tag = $"ID={ID.Value} ";
            else
                args.Tag = "";
            args.Tag += args.FunctionName;

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
            if (functionInfo.CustomAttributes.OfType<LoggingAttribute>().Any())
            {
                var loggingAtt = functionInfo.CustomAttributes.OfType<LoggingAttribute>().First();
                return new FunctionLoggingHandler { ID = loggingAtt.ID };
            }

            return new FunctionLoggingHandler();
        }
    }
}
