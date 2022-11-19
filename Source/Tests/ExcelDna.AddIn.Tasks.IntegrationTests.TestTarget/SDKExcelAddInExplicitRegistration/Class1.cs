using ExcelDna.Integration;

namespace SDKExcelAddInExplicitRegistration
{
    public class Class1
    {
        public static string MyFunction()
        {
            return "MyFunction";
        }

        [ExcelFunction]
        public static string MyExplicitFunction()
        {
            return "MyExplicitFunction";
        }
    }
}
