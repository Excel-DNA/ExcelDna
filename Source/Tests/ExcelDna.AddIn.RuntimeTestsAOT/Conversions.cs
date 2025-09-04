using ExcelDna.Integration;

namespace ExcelDna.AddIn.RuntimeTestsAOT
{
    public class Conversions
    {
        [ExcelParameterConversion]
        public static Version ToVersion(string s)
        {
            return new Version(s);
        }

        [ExcelReturnConversion]
        public static string FromTestType1(TestType1 value)
        {
            return value.Value;
        }
    }
}
