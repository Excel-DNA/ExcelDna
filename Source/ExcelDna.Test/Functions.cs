using ExcelDna.Integration;
using ExcelDna.Registration;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelDna.Test
{
    public static class Functions
    {
        [ExcelFunction("SayHelloFromTest")]
        public static string SayHello(string name) => $"Hello {name}";

        [ExcelAsyncFunction]
        public static async Task<string> MyAsyncTaskHello(string name, int msDelay)
        {
            await Task.Delay(msDelay);
            return $"Hello async task {name}";
        }

        public static object SayHelloSlow(string name)
         => ExcelAsyncUtil.Run(nameof(SayHelloSlow), name, () =>
         {
             Thread.Sleep(4000);
             return "Done " + name;
         });
    }
}
