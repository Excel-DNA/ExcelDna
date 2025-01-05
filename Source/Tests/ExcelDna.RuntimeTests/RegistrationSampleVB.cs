using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
#if DEBUG
    public class RegistrationSampleVB
    {
        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RegistrationSampleVB\bin\Debug\net6.0-windows\ExcelDna.AddIn.RegistrationSampleVB-AddIn")]
        public void Optional()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=dnaOptionalAnswer()";
                Assert.Equal("The answer is 42", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange.Formula = "=dnaOptionalAnswer(7.5)";
                Assert.Equal("The answer is 7.5", functionRange.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RegistrationSampleVB\bin\Debug\net6.0-windows\ExcelDna.AddIn.RegistrationSampleVB-AddIn")]
        public void Params()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=dnaAddValues(3,6,12)";
                Assert.Equal(21, functionRange.Value);
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange.Formula = "=dnaConcatStrings(\"p\",\"-\",\"a\",\"b\")";
                Assert.Equal("pa-b", functionRange.Value.ToString());
            }
        }
    }
#endif
}
