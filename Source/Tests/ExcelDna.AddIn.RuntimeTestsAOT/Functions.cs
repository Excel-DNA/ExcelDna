using ExcelDna.Integration;
using ExcelDna.Registration;

namespace ExcelDna.AddIn.RuntimeTestsAOT
{
    public class Functions
    {
        [ExcelFunction]
        public static string NativeHello(string name)
        {
            return $"Hello {name}!";
        }

        [ExcelFunction]
        public static int NativeSum(int i1, int i2)
        {
            return i1 + i2;
        }

        [ExcelAsyncFunction]
        public static async Task<string> NativeAsyncTaskHello(string name, int msDelay)
        {
            await Task.Delay(msDelay);
            return $"Hello native async task {name}";
        }
    }
}
