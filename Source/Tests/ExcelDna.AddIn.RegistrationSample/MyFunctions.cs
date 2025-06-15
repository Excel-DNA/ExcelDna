using ExcelDna.Integration;
using ExcelDna.Registration;

namespace ExcelDna.AddIn.RegistrationSample
{
    public class MyFunctions
    {
        // Will not be registered in Excel by Excel-DNA, without being picked up by our Registration processing
        // since there is no ExcelFunction attribute, and ExplicitRegistration="true" in the .dna file prevents this 
        // function from being registered by the default processing.
        public static string dnaSayHello(string name)
        {
            return "Hello " + name + "!";
        }

        // Will be picked up by our explicit processing, no conversions applied, and normal registration
        [ExcelFunction(Name = "dnaSayHello")]
        public static string dnaSayHello2(string name)
        {
            if (name == "Bang!") throw new ArgumentException("Bad name!");
            return "Hello " + name + "!";
        }

        [ExcelFunction]
        [Timing]
        public static string dnaSayHelloTiming(string name)
        {
            return "Hello " + name + "!";
        }

        [ExcelFunction]
        [Cache(99)]
        public static string dnaSayHelloCache(string name)
        {
            return "Hello " + name + "!";
        }

        [ExcelFunction]
        [SuppressInDialog]
        public static string dnaSayHelloSuppressInDialog(string name)
        {
            return "Hello " + name + "!";
        }

        [ExcelFunction]
        public static string MyRegistrationSampleFunctionExecutionLog()
        {
            string result = Logger.GetLog();
            Logger.ClearLog();
            return result;
        }
    }
}
