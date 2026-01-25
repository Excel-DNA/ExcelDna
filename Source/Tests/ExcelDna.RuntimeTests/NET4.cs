using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
    public class NET4
    {
        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsNET4)]
        public void OptionalDateTime()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=Net4OptionalDateTime(\"2024/11/21\")";
            Assert.Equal(".NET 4 Optional DateTime: 11/21/2024 12:00:00 AM", functionRange.Value.ToString());

            functionRange.Formula = "=Net4OptionalDateTime()";
            Assert.Equal(".NET 4 Optional DateTime: 1/1/0001 12:00:00 AM", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsNET4)]
        public void DefaultDateTime()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
            functionRange.Formula = "=Net4DateTimeDefault()";
            Assert.Equal<double>(0, functionRange.Value2);

            functionRange.Formula = "=Net4DateTimeDefault(\"2025-10-13\")";
            Assert.Equal<double>(45943, functionRange.Value2);
        }
    }
}
