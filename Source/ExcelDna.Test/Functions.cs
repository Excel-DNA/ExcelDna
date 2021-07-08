using ExcelDna.Integration;
using System;

namespace ExcelDna.Test
{
    public static class Functions
    {
        [ExcelFunction("SayHelloFromTest")]
        public static string SayHello(string name) => $"Hello {name}";
    }
}
