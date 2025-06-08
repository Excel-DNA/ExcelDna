using ExcelDna.Integration;
using ExcelDna.Registration;

namespace ExcelDna.AddIn.RuntimeTestsAOT
{
    public class Functions
    {
        [ExcelFunction]
        public static string NativeHello(string name)
        {
            return $"Hello {name}!";
        }

        [ExcelFunction]
        public static int NativeSum(int i1, int i2)
        {
            return i1 + i2;
        }

        [ExcelAsyncFunction]
        public static async Task<string> NativeAsyncTaskHello(string name, int msDelay)
        {
            await Task.Delay(msDelay);
            return $"Hello native async task {name}";
        }

        [ExcelFunction]
        public static string NativeApplicationName()
        {
            return (string)ExcelDnaUtil.DynamicApplication.GetProperty("Name");
        }

        [ExcelFunction]
        public static double NativeApplicationGetCellValue(string cell)
        {
            var workbook = (IDynamic)ExcelDnaUtil.DynamicApplication.GetProperty("ActiveWorkbook");
            var sheets = (IDynamic)workbook.GetProperty("Sheets");
            var sheet = (IDynamic)sheets[1];
            var range = (IDynamic)sheet.GetProperty("Range", [cell]);
            return (double)range.GetProperty("Value");
        }
    }
}
