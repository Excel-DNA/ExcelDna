using ExcelDna.Integration;

namespace ExcelDna.AddIn.RuntimeTests
{
    public class Conversions
    {
        [ExcelParameterConversion]
        public static TestType1 Order2ToTestType1(string value)
        {
            return new TestType1(value);
        }

        [ExcelParameterConversion]
        public static TestType2 Order1ToTestType2FromTestType1(TestType1 value)
        {
            return new TestType2("From TestType1 " + value.Value);
        }

        [ExcelParameterConversion]
        public static Version ToVersion(string s)
        {
            return new Version(s);
        }
    }
}
