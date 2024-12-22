using ExcelDna.Integration;
using ExcelDna.Registration;

namespace ExcelDna.AddIn.RegistrationSample
{
    public class MyFunctions
    {
        [ExcelFunction]
        public static string MyRegistrationSampleFunctionExecutionLog()
        {
            string result = Logger.GetLog();
            Logger.ClearLog();
            return result;
        }
    }
}
