using ExcelDna.Integration;
using System;
using System.Threading;

namespace ExcelDna.Test
{
    public static class Functions
    {
        [ExcelFunction("SayHelloFromTest")]
        public static string SayHello(string name) => $"Hello {name}";

        public static object SayHelloSlow(string name)
         => ExcelAsyncUtil.Run(nameof(SayHelloSlow), name, () =>
         {
             Thread.Sleep(4000);
             return "Done " + name;
         });
    }
}
