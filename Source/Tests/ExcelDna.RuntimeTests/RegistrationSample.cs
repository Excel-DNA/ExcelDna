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
            {
                functionRange.Formula = "=MyRegistrationSampleFunctionExecutionLog()";
                functionRange.Formula = "=dnaSayHello(\"FunctionExecutionHandler\")";
                functionRange.Formula = "=MyRegistrationSampleFunctionExecutionLog()";

                Assert.True(functionRange.Value.ToString().Contains("FunctionLoggingHandler dnaSayHello - OnEntry - Args: FunctionExecutionHandler"));
            }
            {
                functionRange.Formula = "=MyRegistrationSampleFunctionExecutionLog()";
                functionRange.Formula = "=dnaSayHelloTiming()";
                functionRange.Formula = "=MyRegistrationSampleFunctionExecutionLog()";

                Assert.True(functionRange.Value.ToString().Contains("TimingFunctionExecutionHandler dnaSayHelloTiming"));
            }
            {
                functionRange.Formula = "=MyRegistrationSampleFunctionExecutionLog()";

                functionRange.Formula = "=dnaSayHelloCache(\"123\")";
                functionRange.Formula = "=MyRegistrationSampleFunctionExecutionLog()";
                Assert.True(functionRange.Value.ToString().Contains("CacheFunctionExecutionHandler dnaSayHelloCache result not in cache"));

                functionRange.Formula = "=dnaSayHelloCache(\"123\")";
                functionRange.Formula = "=MyRegistrationSampleFunctionExecutionLog()";
                Assert.True(functionRange.Value.ToString().Contains("CacheFunctionExecutionHandler dnaSayHelloCache result in cache"));
            }
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RegistrationSample\bin\Debug\net6.0-windows\ExcelDna.AddIn.RegistrationSample-AddIn")]
        public void InstanceMemberRegistration()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=GetContent(\"world\")";
                Assert.Equal("Content is world", functionRange.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RegistrationSample\bin\Debug\net6.0-windows\ExcelDna.AddIn.RegistrationSample-AddIn")]
        public void ParameterConversion()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=dnaConversionTest(B2)";
                Assert.True(functionRange.Value.ToString().StartsWith("Reference: "));
                Assert.True(functionRange.Value.ToString().EndsWith("!$B$2"));
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRange1.Value = "Hello There!";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D2"];
                functionRange2.Formula = "=dnaConversionToString(D1)";
                Assert.Equal("Hello There!", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D3"];
                functionRange3.Formula = "=dnaDirectString(D1)";
                Assert.Equal("Hello There!", functionRange3.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange1.Value = "3.5";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRange2.Formula = "=dnaConversionToDouble(C1)";
                Assert.Equal("3.5", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C3"];
                functionRange3.Formula = "=dnaDirectDouble(C1)";
                Assert.Equal("3.5", functionRange3.Value.ToString());

                Range functionRange4 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C4"];
                functionRange4.Formula = "=dnaDirectDouble()";
                Assert.Equal("0", functionRange4.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E1"];
                functionRange1.Value = "3.5";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E2"];
                functionRange2.Formula = "=dnaConversionToInt32(E1)";
                Assert.Equal("4", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E3"];
                functionRange3.Formula = "=dnaDirectInt32(E1)";
                Assert.Equal("4", functionRange3.Value.ToString());

                Range functionRange4 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E4"];
                functionRange4.Formula = "=dnaDirectInt32()";
                Assert.Equal("0", functionRange4.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E11"];
                functionRange1.Value = "1";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E12"];
                functionRange2.Formula = "=dnaConversionToInt32(E11)";
                Assert.Equal("1", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E13"];
                functionRange3.Formula = "=dnaDirectInt32(E11)";
                Assert.Equal("1", functionRange3.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F1"];
                functionRange1.Value = "3.5";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F2"];
                functionRange2.Formula = "=dnaConversionToInt64(F1)";
                Assert.Equal("4", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F3"];
                functionRange3.Formula = "=dnaDirectInt64(F1)";
                Assert.Equal("4", functionRange3.Value.ToString());

                Range functionRange4 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F4"];
                functionRange4.Formula = "=dnaDirectInt64()";
                Assert.Equal("0", functionRange4.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F11"];
                functionRange1.Value = "1";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F12"];
                functionRange2.Formula = "=dnaConversionToInt64(F11)";
                Assert.Equal("1", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F13"];
                functionRange3.Formula = "=dnaDirectInt64(F11)";
                Assert.Equal("1", functionRange3.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["G1"];
                functionRange1.Value = "12/18/1979  1:12:00 PM";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["G2"];
                functionRange2.Formula = "=dnaConversionToDateTime(G1)";
                Assert.Equal("29207.55", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["G3"];
                functionRange3.Formula = "=dnaDirectDateTime(G1)";
                Assert.Equal("29207.55", functionRange3.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["H1"];
                functionRange1.Value = "TRUE";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["H2"];
                functionRange2.Formula = "=dnaConversionToBoolean(H1)";
                Assert.Equal("True", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["H3"];
                functionRange3.Formula = "=dnaDirectBoolean(H1)";
                Assert.Equal("True", functionRange3.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["H11"];
                functionRange1.Value = "FALSE";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["H12"];
                functionRange2.Formula = "=dnaConversionToBoolean(H11)";
                Assert.Equal("False", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["H13"];
                functionRange3.Formula = "=dnaDirectBoolean(H11)";
                Assert.Equal("False", functionRange3.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RegistrationSample\bin\Debug\net6.0-windows\ExcelDna.AddIn.RegistrationSample-AddIn")]
        public void ParameterConversionNullableOptional()
        {
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange1.Formula = "=dnaParameterConvertTest()";
                Assert.Equal("NULL!!!", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRange2.Formula = "=dnaParameterConvertTest(2.5)";
                Assert.Equal("2.5", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange1.Formula = "=dnaDoubleNullableOptional()";
                Assert.Equal("NaN", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRange2.Formula = "=dnaDoubleNullableOptional(2.5)";
                Assert.Equal("2.5", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRange1.Formula = "=dnaParameterConvertOptionalTest()";
                Assert.Equal("VALUE: 42.0", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D2"];
                functionRange2.Formula = "=dnaParameterConvertOptionalTest(2.5)";
                Assert.Equal("VALUE: 2.5", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E1"];
                functionRange1.Formula = "=dnaMultipleOptional()";
                Assert.Equal("VALUES: 3.1415927 & @42@", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E2"];
                functionRange2.Formula = "=dnaMultipleOptional(2.5)";
                Assert.Equal("VALUES: 2.5000000 & @42@", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E3"];
                functionRange3.Formula = "=dnaMultipleOptional(2.5, \"abc\")";
                Assert.Equal("VALUES: 2.5000000 & abc", functionRange3.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F1"];
                functionRange1.Formula = "=dnaOptionalInt()";
                Assert.Equal("VALUE: 42.0", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F2"];
                functionRange2.Formula = "=dnaOptionalInt(2)";
                Assert.Equal("VALUE: 2.0", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["G1"];
                functionRange1.Formula = "=dnaOptionalString()";
                Assert.Equal("Hello World!", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["G2"];
                functionRange2.Formula = "=dnaOptionalString(\"abc\")";
                Assert.Equal("abc", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["H1"];
                functionRange1.Formula = "=dnaNullableOptionalDateTime()";
                Assert.Equal("NULL", functionRange1.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RegistrationSample\bin\Debug\net6.0-windows\ExcelDna.AddIn.RegistrationSample-AddIn")]
        public void ParameterConversionNullable()
        {
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange1.Formula = "=dnaNullableDouble()";
                Assert.Equal("NULL", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRange2.Formula = "=dnaNullableDouble(12.3)";
                Assert.Equal("VAL: 12.3", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange1.Formula = "=dnaNullableInt()";
                Assert.Equal("NULL", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRange2.Formula = "=dnaNullableInt(12)";
                Assert.Equal("VAL: 12", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRange1.Formula = "=dnaNullableLong()";
                Assert.Equal("NULL", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D2"];
                functionRange2.Formula = "=dnaNullableLong(12345)";
                Assert.Equal("VAL: 12345", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E1"];
                functionRange1.Formula = "=dnaNullableDateTime()";
                Assert.Equal("NULL", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E2"];
                functionRange2.Formula = "=dnaNullableDateTime(12.3)";
                Assert.Equal("VAL: 1/11/1900 7:12:00 AM", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F1"];
                functionRange1.Formula = "=dnaNullableBoolean()";
                Assert.Equal("NULL", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F2"];
                functionRange2.Formula = "=dnaNullableBoolean(2)";
                Assert.Equal("VAL: True", functionRange2.Value.ToString());
            }
        }
    }
#endif
}
