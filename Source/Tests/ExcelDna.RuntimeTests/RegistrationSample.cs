using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
#if DEBUG
    public class RegistrationSample
    {
        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RegistrationSample\bin\Debug\net6.0-windows\ExcelDna.AddIn.RegistrationSample-AddIn")]
        public void SayHello()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=dnaSayHello(\"world\")";
                Assert.Equal("Hello world!", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange.Formula = "=dnaSayHello(\"Bang!\")";
                Assert.True(functionRange.Value.ToString().StartsWith("!!! ERROR: System.ArgumentException: Bad name!"));
            }
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RegistrationSample\bin\Debug\net6.0-windows\ExcelDna.AddIn.RegistrationSample-AddIn")]
        public void FunctionExecutionHandler()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];

            functionRange.Formula = "=MyRegistrationSampleFunctionExecutionLog()";
            functionRange.Formula = "=dnaSayHello(\"FunctionExecutionHandler\")";
            functionRange.Formula = "=MyRegistrationSampleFunctionExecutionLog()";

            Assert.True(functionRange.Value.ToString().Contains("FunctionLoggingHandler dnaSayHello - OnEntry - Args: FunctionExecutionHandler"));
        }
    }
#endif
}
