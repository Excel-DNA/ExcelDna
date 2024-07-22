using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
#if DEBUG
    public class UserCodeConversions
    {
        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void ReferenceToRange()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=MyUserCodeConversionReferenceToRange(B2)";
                Assert.Equal("$B$2", functionRange.Value.ToString());
            }
        }
    }
#endif
}
