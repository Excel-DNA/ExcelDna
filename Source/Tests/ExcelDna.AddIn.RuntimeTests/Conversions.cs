using ExcelDna.Integration;

namespace ExcelDna.AddIn.RuntimeTests
{
    public class Conversions
    {
        [ExcelParameterConversion]
        public static TestType1 ToTestType1(string value)
        {
            return new TestType1(value);
        }

        [ExcelParameterConversion]
        public static Version ToVersion(string s)
        {
            return new Version(s);
        }
    }
}
