using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
#if DEBUG
    [ExcelTestSettings(OutOfProcess = true)]
    [Collection("OutOfProcess")]
    public class ObjectHandleOutOfProcess
    {
        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void ThreadSafe()
        {
            for (int i = 1; i <= 5; ++i)
            {
                {
                    Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range[$"C{i}"];
                    functionRange.Formula = $"=MyCreateCalcTS({(i - 1) * 20}, 0)";
                }
                {
                    Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range[$"D{i}"];
                    functionRange.Formula = $"=MyCalcSumTS(C{i})";
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
                Range functionRangeC1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRangeC1.Formula = "=MyGetCreatedDisposableObjectsCount()";
                int initialCreatedObjectsCount = (int)functionRangeC1.Value;

                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange1.Formula = "=MyCreateDisposableObject(1)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRange2.Formula = "=MyGetDisposableObjectsCount()";

                Range functionRangeC2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRangeC2.Formula = "=MyGetCreatedDisposableObjectsCount()";
                int finalCreatedObjectsCount = (int)functionRangeC2.Value;

                Assert.Equal(1, finalCreatedObjectsCount - initialCreatedObjectsCount);

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
