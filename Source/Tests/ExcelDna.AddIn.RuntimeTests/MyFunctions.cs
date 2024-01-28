using ExcelDna.Integration;

namespace ExcelDna.AddIn.RuntimeTests
{
    public class MyFunctions
    {
        [ExcelFunction]
        public static string SayHello(string name)
        {
            return $"Hello {name}";
        }
    }
}
