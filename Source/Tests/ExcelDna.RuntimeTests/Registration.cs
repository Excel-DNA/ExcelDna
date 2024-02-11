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
#endif
    }
}
