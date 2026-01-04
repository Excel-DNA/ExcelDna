using ExcelDna.Registration;

namespace ExcelDna.TestAOT
{
    public class Functions
    {
        [ExcelAsyncFunction]
        public static async Task<string> NativeAsyncTaskHello(string name, int msDelay)
        {
            await Task.Delay(msDelay);
            return $"Hello native async task {name}";
        }
    }
}
