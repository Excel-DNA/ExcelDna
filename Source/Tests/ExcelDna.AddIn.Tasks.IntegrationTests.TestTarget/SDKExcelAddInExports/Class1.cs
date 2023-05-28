using ExcelDna.Integration;

namespace SDKExcelAddInExports
{
    public class Class1
    {
        [ExcelFunction]
        public static string MyMainFunction()
        {
            return "MyMainFunction";
        }
    }
}
