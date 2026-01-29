using ExcelDna.Integration;
using System;

namespace ExcelDna.AddIn.RuntimeTestsNET4
{
    public class Functions
    {
        [ExcelFunction]
        public static string Net4OptionalDateTime(DateTime dt = default)
        {
            return ".NET 4 Optional DateTime: " + dt.ToString();
        }

        [ExcelFunction]
        public static double Net4DateTimeDefault(DateTime value = default) => value.ToOADate();
    }
}
