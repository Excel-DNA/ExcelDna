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

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void Disposable()
        {
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange1.Formula = "=MyCreateDisposableObject(1)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRange2.Formula = "=MyGetDisposableObjectsCount()";

                Assert.Equal("1", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange1.Formula = "=MyCreateDisposableObject(5)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRange2.Formula = "=MyGetDisposableObjectsCount()";

                Assert.Equal("2", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange1.Clear();

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRange2.Formula = "=MyGetDisposableObjectsCount()";

                Assert.Equal("1", functionRange2.Value.ToString());
            }
        }
    }
#endif
}
