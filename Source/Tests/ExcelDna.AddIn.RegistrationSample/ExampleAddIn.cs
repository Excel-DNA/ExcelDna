using ExcelDna.Integration;
using ExcelDna.Registration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDna.AddIn.RegistrationSample
{
    public class ExampleAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex => "!!! ERROR: " + ex.ToString());

            var functionHandlerConfig = GetFunctionExecutionHandlerConfig();

            ExcelRegistration.GetExcelFunctions()
                .ProcessAsyncRegistrations(nativeAsyncIfAvailable: false)
                .ProcessFunctionExecutionHandlers(functionHandlerConfig)
                .RegisterFunctions()
                ;

            // First example if Instance -> Static conversion
            InstanceMemberRegistration.TestInstanceRegistration();
        }

        public void AutoClose()
        {
        }

        static FunctionExecutionConfiguration GetFunctionExecutionHandlerConfig()
        {
            return new FunctionExecutionConfiguration()
                .AddFunctionExecutionHandler(FunctionLoggingHandler.LoggingHandlerSelector)
                .AddFunctionExecutionHandler(CacheFunctionExecutionHandler.CacheHandlerSelector)
                .AddFunctionExecutionHandler(TimingFunctionExecutionHandler.TimingHandlerSelector)
                .AddFunctionExecutionHandler(SuppressInDialogFunctionExecutionHandler.SuppressInDialogSelector);
        }
    }
}
