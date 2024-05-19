using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
#if DEBUG
    [ExcelTestSettings(OutOfProcess = true)]
    public class ObjectHandleOutOfProcess
    {
        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void ThreadSafe()
        {
            for (int i = 1; i <= 5; ++i)
            {
                {
                    Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range[$"C{i}"];
                    functionRange.Formula = $"=MyCreateObjectTS({(i - 1) * 20})";
                }
                {
                    Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range[$"D{i}"];
                    functionRange.Formula = $"=MyUseObjectTS(C{i})";
                }
            }

            for (int i = 1; i <= 5; ++i)
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range[$"D{i}"];
                Assert.Equal($"{(i - 1) * 20}", functionRange.Value.ToString());
            }
        }
    }
#endif
}
