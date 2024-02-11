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

        [ExcelFunction]
        public static string MyDouble(double d)
        {
            return d.ToString();
        }

        [ExcelFunction]
        public static string MyNullableDouble(double? d)
        {
            return "Nullable VAL: " + (d.HasValue ? d : "NULL");
        }

        [ExcelFunction]
        public static string MyOptionalDouble(double d = 1.23)
        {
            return "Optional VAL: " + d.ToString();
        }

        [ExcelFunction]
        public static string MyEnum(DateTimeKind e)
        {
            return "Enum VAL: " + e.ToString();
        }

        [ExcelMapArrayFunction]
        public static IEnumerable<string> MyMapArray(IEnumerable<DateTimeKind> a)
        {
            return a.Select(i => "Array element VAL: " + i.ToString());
        }
    }
}
