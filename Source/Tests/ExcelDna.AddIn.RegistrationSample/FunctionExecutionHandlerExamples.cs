using ExcelDna.Integration;
using ExcelDna.Registration;

namespace ExcelDna.AddIn.RegistrationSample
{
    public static class FunctionExecutionHandlerExamples
    {
        // Slow function that will be cached for 30 seconds
        // F2, Enter or Ctrl+Alt+F9 should be fast for a while, then slow if you wait or change inputs
        // Timings will be writted to the Debug Output.
        [ExcelFunction, Cache(30), Timing, SuppressInDialog]
        public static string dnaSleepFirstTime(string input)
        {
            System.Threading.Thread.Sleep(5000);
            return input;
        }

        [ExcelFunction, Timing]
        public static double dnaCountUpTo(long bigNumber = 1000000)
        {
            double total = 0;
            for (long l = 0; l < bigNumber; l++)
            {
                total += l;
            }
            return total;
        }
    }
}
