using ExcelDna.Integration;

namespace ExcelDna.Test
{
    public class Functions
    {
        [ExcelFunction]
        public static string MyHello(string name)
        {
            return $"[ExcelDna.Test] Hello {name}!";
        }

        [ExcelFunction]
        public static int MySum(int i1, int i2)
        {
            return i1 + i2;
        }
    }
}
