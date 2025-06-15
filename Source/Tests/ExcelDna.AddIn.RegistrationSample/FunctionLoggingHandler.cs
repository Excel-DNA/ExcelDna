using System.Diagnostics;
using System.Linq;
using ExcelDna.Registration;

namespace ExcelDna.AddIn.RegistrationSample
{
    public class FunctionLoggingHandler : FunctionExecutionHandler
    {
        int Index;
        public override void OnEntry(FunctionExecutionArgs args)
        {
            // FunctionExecutionArgs gives access to the function name and parameters,
            // and gives some options for flow redirection.

            // Tag will flow through the whole handler
            args.Tag = args.FunctionName + ":" + Index;
            string arguments = string.Join(",", args.Arguments.Select(arg => arg.ToString()));
            Debug.Print("FunctionLoggingHandler {0} - OnEntry - Args: {1}", args.Tag, arguments);
            Logger.Log($"FunctionLoggingHandler {args.FunctionName} - OnEntry - Args: {arguments}");
        }

        public override void OnSuccess(FunctionExecutionArgs args)
        {
            Debug.Print("FunctionLoggingHandler {0} - OnSuccess - Result: {1}", args.Tag, args.ReturnValue);
        }

        public override void OnException(FunctionExecutionArgs args)
        {
            Debug.Print("FunctionLoggingHandler {0} - OnException - Message: {1}", args.Tag, args.Exception);
        }

        public override void OnExit(FunctionExecutionArgs args)
        {
            Debug.Print("FunctionLoggingHandler {0} - OnExit", args.Tag);
            Logger.Log($"FunctionLoggingHandler {args.Tag} - OnExit");
        }

        // The configuration part - maybe move somewhere else.
        // (Add a registration index just to show we can attach arbitrary data to the captured handler instance which may be created for each function.)
        // If we return the same object for every function, the object needs to be re-entrancy safe is used by IsThreadSafe functions.
        static int _index = 0;
        internal static FunctionExecutionHandler LoggingHandlerSelector(ExcelFunctionRegistration functionRegistration)
        {
            return new FunctionLoggingHandler { Index = _index++ };
        }
    }
}
