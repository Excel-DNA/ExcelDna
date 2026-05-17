using Microsoft.Office.Interop.Excel;
using System.IO;
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
            DateAssertions.EqualPrefixed(functionRange.Value, ".NET 4 Optional DateTime: ", new DateTime(2024, 11, 21));

            functionRange.Formula = "=Net4OptionalDateTime()";
            DateAssertions.EqualPrefixed(functionRange.Value, ".NET 4 Optional DateTime: ", DateTime.MinValue);
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

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsNET4)]
        public void BaseDirectory()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
            functionRange.Formula = "=Net4BaseDirectory()";
            Assert.EndsWith(Path.GetFileName(Path.GetDirectoryName(AddInPath.RuntimeTestsNET4)), functionRange.Value2);
        }
    }
}
