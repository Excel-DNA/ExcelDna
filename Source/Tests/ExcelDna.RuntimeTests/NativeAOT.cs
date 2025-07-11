﻿using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
    public class NativeAOT
    {
        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsAOT)]
        public void Hello()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=NativeHello(\"world\")";

            Assert.Equal("Hello world!", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsAOT)]
        public void Sum()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=NativeSum(3, 4)";

            Assert.Equal("7", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsAOT)]
        public void AsyncTaskInstant()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=NativeAsyncTaskHello(\"world\", 0)";

            Assert.Equal("Hello native async task world", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsAOT)]
        public void DefaultAsyncReturnValue()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=NativeAsyncTaskHello(\"world\", 1000)";

            Assert.Equal(-2146826246, functionRange.Value); // #N/A
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsAOT)]
        public void DynamicApplication()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
                functionRange.Formula = "=NativeApplicationName()";

                Assert.Equal("Microsoft Excel", functionRange.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange1.Value = 4.2;

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRange2.Formula = "=NativeApplicationGetCellValue(\"C1\")";

                Assert.Equal(4.2, functionRange2.Value);
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRange1.Value = 41.22;

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D2"];
                functionRange2.Formula = "=NativeApplicationGetCellValueT(\"D1\")";

                Assert.Equal(41.22, functionRange2.Value);
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E1"];

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E2"];
                functionRange2.Formula = "=NativeApplicationAddCellComment(\"E1\", \"Native Comment\")";

                Assert.Equal("Native Comment", functionRange2.Value);
                Assert.Equal("Native Comment", functionRange1.Comment.Text());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsAOT)]
        public void RangeArray2D()
        {
            Range functionRangeA1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A1"];
            functionRangeA1.Value = "str";

            Range functionRangeA2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A2"];
            functionRangeA2.Value = 1;

            Range functionRangeB1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
            functionRangeB1.Value = true;

            Range functionRangeB2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
            functionRangeB2.Value = 3.5;

            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
            functionRange.Formula = "=NativeRangeConcat2(A1:B2)";

            Assert.Equal("strTrue13.5", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsAOT)]
        public void Nullable()
        {
            Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A1"];
            functionRange1.Formula = "=NativeNullableDouble(1.2)";
            Assert.Equal("Native Nullable VAL: 1.2", functionRange1.Value.ToString());

            Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A2"];
            functionRange2.Formula = "=NativeNullableDouble()";
            Assert.Equal("Native Nullable VAL: NULL", functionRange2.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsAOT)]
        public void Optional()
        {
            Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A1"];
            functionRange1.Formula = "=NativeOptionalDouble(2.3)";
            Assert.Equal("Native Optional VAL: 2.3", functionRange1.Value.ToString());

            Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A2"];
            functionRange2.Formula = "=NativeOptionalDouble()";
            Assert.Equal("Native Optional VAL: 1.23", functionRange2.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsAOT)]
        public void Range()
        {
            Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A1"];
            functionRange1.Formula = "=NativeRangeAddress(B2)";
            Assert.Equal("Native Address: $B$2", functionRange1.Value.ToString());

            Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A2"];
            functionRange2.Formula = "=NativeRangeAddress(B2:C4)";
            Assert.Equal("Native Address: $B$2:$C$4", functionRange2.Value.ToString());

            Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A3"];
            functionRange3.Formula = "=NativeRangeAddress((B2,D5:E6))";
            Assert.Equal("Native Address: $B$2,$D$5:$E$6", functionRange3.Value.ToString());
        }
    }
}
