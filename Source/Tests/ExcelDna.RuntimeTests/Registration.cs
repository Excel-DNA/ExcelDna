using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
#if DEBUG
    public class Registration
    {
        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void SayHello()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=SayHello(\"world\")";
            Assert.Equal("Hello world", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void Double()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyDouble(3.4)";
            Assert.Equal("3.4", functionRange.Value.ToString());

            functionRange.Formula = "=MyDouble()";
            Assert.Equal("0", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void NullableDouble()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyNullableDouble(1.2)";
            Assert.Equal("Nullable VAL: 1.2", functionRange.Value.ToString());

            functionRange.Formula = "=MyNullableDouble()";
            Assert.Equal("Nullable VAL: NULL", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void OptionalDouble()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyOptionalDouble(2.3)";
            Assert.Equal("Optional VAL: 2.3", functionRange.Value.ToString());

            functionRange.Formula = "=MyOptionalDouble()";
            Assert.Equal("Optional VAL: 1.23", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void Enum()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyEnum(\"Unspecified\")";
            Assert.Equal("Enum VAL: Unspecified", functionRange.Value.ToString());

            functionRange.Formula = "=MyEnum(\"Local\")";
            Assert.Equal("Enum VAL: Local", functionRange.Value.ToString());

            functionRange.Formula = "=MyEnum(1)";
            Assert.Equal("Enum VAL: Utc", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void EnumReturn()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyEnumReturn(\"Unspecified\")";
            Assert.Equal("Unspecified", functionRange.Value.ToString());

            functionRange.Formula = "=MyEnumReturn(\"Local\")";
            Assert.Equal("Local", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void MapArray()
        {
            Range a1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A1:A1"];
            a1.Value = "Utc";

            Range a2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A2:A2"];
            a2.Value = "Local";

            Range a3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A3:A3"];
            a3.Value = "Unspecified";

            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B3"];
            functionRange.FormulaArray = "=MyMapArray(A1:A3)";

            Range b1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            Assert.Equal("Array element VAL: Utc", b1.Value.ToString());

            Range b2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2:B2"];
            Assert.Equal("Array element VAL: Local", b2.Value.ToString());

            Range b3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B3:B3"];
            Assert.Equal("Array element VAL: Unspecified", b3.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void AsyncTaskInstant()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyAsyncTaskHello(\"world\", 0)";

            Assert.Equal("Hello async task world", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void TaskInstant()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyTaskHello(\"world\")";

            Assert.Equal("Hello task world", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void StringArray()
        {
            Range a1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A1:A1"];
            a1.Value = "01";

            Range a2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A2:A2"];
            a2.Value = "2.30";

            Range a3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A3:A3"];
            a3.Value = "World";

            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyStringArray(A1:A3)";

            Assert.Equal("StringArray VALS: 12.3World", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void StringArray2D()
        {
            Range a1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A1"];
            a1.Value = "01";

            Range a2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A2"];
            a2.Value = "2.30";

            Range a3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A3"];
            a3.Value = "Hello";

            Range b1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
            b1.Value = "5";

            Range b2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
            b2.Value = "6.7";

            Range b3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B3"];
            b3.Value = "World";

            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
            functionRange.Formula = "=MyStringArray2D(A1:B3)";

            Assert.Equal("StringArray2D VALS: 15 2.36.7 HelloWorld ", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void UserDefinedParameterConversions()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];

            functionRange.Formula = "=MyVersion2(\"4.3.2.1\")";
            Assert.Equal("The Version value with field count 2 is 4.3", functionRange.Value.ToString());

            functionRange.Formula = "=MyTestType1(\"world\")";
            Assert.Equal("The TestType1 value is world", functionRange.Value.ToString());

            functionRange.Formula = "=MyTestType2(\"world2\")";
            Assert.Equal("The TestType2 value is From TestType1 world2", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void FunctionExecutionHandlerExtended()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];

            functionRange.Formula = "=MyVersion2(\"5.4.3.2\")";
            functionRange.Formula = "=MyFunctionExecutionLog()";
            Assert.True(functionRange.Value.ToString().Contains("MyVersion2 - OnSuccess - Result: The Version value with field count 2 is 5.4"));
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void FunctionExecutionHandlerStandard()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];

            functionRange.Formula = "=SayHello(\"FunctionExecutionHandlerStandard\")";
            functionRange.Formula = "=MyFunctionExecutionLog()";
            Assert.True(functionRange.Value.ToString().Contains("SayHello - OnSuccess - Result: Hello FunctionExecutionHandlerStandard"));
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void FunctionExecutionHandlerWithAttribute()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];

            functionRange.Formula = "=SayHelloWithLoggingID(\"FunctionExecutionHandlerWithAttribute\")";
            functionRange.Formula = "=MyFunctionExecutionLog()";
            Assert.True(functionRange.Value.ToString().Contains("ID=7 SayHelloWithLoggingID - OnSuccess - Result: Hello FunctionExecutionHandlerWithAttribute"));
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void ObjectHandles()
        {
            string b1;
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange1.Formula = "=MyCreateObject(45)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRange2.Formula = "=MyUseObject(B1)";

                b1 = functionRange1.Value.ToString();
                Assert.Equal("45", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange1.Formula = "=MyCreateObject(45)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRange2.Formula = "=MyUseObject(C1)";

                Assert.Equal(b1, functionRange1.Value.ToString());
                Assert.Equal("45", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRange1.Formula = "=MyCreateObject(54)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D2"];
                functionRange2.Formula = "=MyUseObject(D1)";

                Assert.NotEqual(b1, functionRange1.Value.ToString());
                Assert.Equal("54", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E1"];
                functionRange1.Formula = "=MyCreateObject2(45)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E2"];
                functionRange2.Formula = "=MyUseObject(E1)";

                Assert.NotEqual(b1, functionRange1.Value.ToString());
                Assert.Equal("90", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F1"];
                functionRange1.Formula = "=MyCreateCalc(1.2, 3.4)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F2"];
                functionRange2.Formula = "=MyCalcSum(F1)";

                Assert.Equal("4.6", functionRange2.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void ObjectHandlesDisposable()
        {
            string b1;
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange1.Formula = "=MyCreateDisposableObject(1)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRange2.Formula = "=MyGetDisposableObjectsCount()";

                b1 = functionRange1.Value.ToString();
                Assert.Equal("1", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange1.Formula = "=MyCreateDisposableObject(5)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRange2.Formula = "=MyGetDisposableObjectsCount()";

                Assert.NotEqual(b1, functionRange1.Value.ToString());
                Assert.Equal("2", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRange1.Formula = "=MyCreateDisposableObject(1)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D2"];
                functionRange2.Formula = "=MyGetDisposableObjectsCount()";

                Assert.Equal(b1, functionRange1.Value.ToString());
                Assert.Equal("2", functionRange2.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void Range()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=MyRange(B2)";
                Assert.Equal("$B$2", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=MyRange(B2:C4)";
                Assert.Equal("$B$2:$C$4", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=MyRange((B2,D5:E6))";
                Assert.Equal("$B$2,$D$5:$E$6", functionRange.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void Params()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=MyParamsFunc1(1,\"2\",4,5)";
                Assert.Equal("1,2, : 2", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRange.Formula = "=MyParamsFunc2(\"a\",,\"c\",\"d\",,\"f\")";
                Assert.Equal("a,,c, [3: d,ExcelDna.Integration.ExcelMissing,f]", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B3"];
                functionRange.Formula = "=MyParamsJoinString(\"//\",\"5\",\"4\",\"3\")";
                Assert.Equal("5//4//3", functionRange.Value.ToString());
            }
        }
    }
#endif
}
