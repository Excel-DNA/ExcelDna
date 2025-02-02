using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using System.Numerics;
using ExcelDna.Registration;

namespace ExcelDna.AddIn.RegistrationSample
{
    // TODO: Fix double->int and similar conversions when objects are received
    //       Test multi-hop return type conversions
    //       Test ExcelArgumentAttributes are preserved.

    public static class ParameterConversionExamples
    {
        // Explore conversions from object -> different types
        [ExcelFunction(IsMacroType = true)]
        public static string dnaConversionTest([ExcelArgument(AllowReference = true)] object arg)
        {
            // This is the full gamut we need to support
            if (arg is double)
                return "Double: " + (double)arg;
            else if (arg is string)
                return "String: " + (string)arg;
            else if (arg is bool)
                return "Boolean: " + (bool)arg;
            else if (arg is ExcelError)
                return "ExcelError: " + arg.ToString();
            else if (arg is object[,])
                // The object array returned here may contain a mixture of different types,
                // reflecting the different cell contents.
                return string.Format("Array[{0},{1}]", ((object[,])arg).GetLength(0), ((object[,])arg).GetLength(1));
            else if (arg is ExcelMissing)
                return "<<Missing>>";
            else if (arg is ExcelEmpty)
                return "<<Empty>>";
            else if (arg is ExcelReference)
                // Calling xlfRefText here requires IsMacroType=true for this function.
                return "Reference: " + XlCall.Excel(XlCall.xlfReftext, arg, true);
            else
                return "!? Unheard Of ?!";
        }

        [Flags]
        internal enum XlType12 : uint
        {
            XlTypeNumber = 0x0001,
            XlTypeString = 0x0002,
            XlTypeBoolean = 0x0004,
            XlTypeReference = 0x0008,
            XlTypeError = 0x0010,
            XlTypeArray = 0x0040,
            XlTypeMissing = 0x0080,
            XlTypeEmpty = 0x0100,
            XlTypeInt = 0x0800,     // int16 in XlOper, int32 in XlOper12, never passed into UDF
        }

        [ExcelFunction]
        public static string dnaConversionToString([ExcelArgument(AllowReference = true)] object arg)
        {
            return (string)XlCall.Excel(XlCall.xlCoerce, arg, (int)XlType12.XlTypeString);
        }

        [ExcelFunction]
        public static string dnaDirectString(string arg)
        {
            return arg;
        }

        [ExcelFunction]
        public static double dnaConversionToDouble([ExcelArgument(AllowReference = true)] object arg)
        {
            return (double)XlCall.Excel(XlCall.xlCoerce, arg, (int)XlType12.XlTypeNumber);
        }

        [ExcelFunction]
        public static double dnaDirectDouble(double arg)
        {
            return arg;
        }

        [ExcelFunction]
        public static int dnaConversionToInt32([ExcelArgument(AllowReference = true)] object arg)
        {
            // The explicit conversion to Int does truncation, which is different to how functions registered
            // as taking an int argument (type "J") are handled.
            object result = XlCall.Excel(XlCall.xlCoerce, arg, (int)XlType12.XlTypeInt);
            double numResult = (double)XlCall.Excel(XlCall.xlCoerce, arg, (int)XlType12.XlTypeNumber);
            return (int)Math.Round(numResult, MidpointRounding.AwayFromZero);
        }

        [ExcelFunction]
        public static int dnaDirectInt32(int arg)
        {
            return arg;
        }

        [ExcelFunction]
        public static long dnaConversionToInt64([ExcelArgument(AllowReference = true)] object arg)
        {
            double numResult = (double)XlCall.Excel(XlCall.xlCoerce, arg, (int)XlType12.XlTypeNumber);
            return (long)Math.Round(numResult, MidpointRounding.AwayFromZero);
        }

        [ExcelFunction]
        public static long dnaDirectInt64(long arg)
        {
            return arg;
        }

        [ExcelFunction]
        public static DateTime dnaConversionToDateTime([ExcelArgument(AllowReference = true)] object arg)
        {
            return DateTime.FromOADate((double)XlCall.Excel(XlCall.xlCoerce, arg, (int)XlType12.XlTypeNumber));
        }

        [ExcelFunction]
        public static DateTime dnaDirectDateTime(DateTime arg)
        {
            return arg;
        }

        [ExcelFunction]
        public static bool dnaConversionToBoolean([ExcelArgument(AllowReference = true)] object arg)
        {
            // This is the full gamut we need to support
            return (bool)XlCall.Excel(XlCall.xlCoerce, arg, (int)XlType12.XlTypeBoolean);
        }

        [ExcelFunction]
        public static bool dnaDirectBoolean(bool arg)
        {
            return arg;
        }

        [ExcelFunction]
        public static string dnaParameterConvertTest(double? optTest)
        {
            if (!optTest.HasValue) return "NULL!!!";

            return optTest.Value.ToString("F1");
        }

        [ExcelFunction]
        public static string dnaDoubleNullableOptional(double? arg = double.NaN)
        {
            return arg.ToString();
        }

        [ExcelFunction]
        public static string dnaParameterConvertOptionalTest(double optOptTest = 42.0)
        {
            return "VALUE: " + optOptTest.ToString("F1");
        }

        [ExcelFunction]
        public static string dnaMultipleOptional(double optOptTest1 = 3.14159265, string optOptTest2 = "@42@")
        {
            return "VALUES: " + optOptTest1.ToString("F7") + " & " + optOptTest2;
        }

        // Problem function
        // This function cannot be called yet, since the cast from what Excel passes in (object) to the int we expect fails.
        // It will need some improved conversion function in the OptionalParameterConversion.
        [ExcelFunction]
        public static string dnaOptionalInt(int optOptTest = 42)
        {
            return "VALUE: " + optOptTest.ToString("F1");
        }

        [ExcelFunction]
        public static string dnaOptionalString(string optOptTest = "Hello World!")
        {
            return optOptTest;
        }

        [ExcelFunction]
        public static string dnaNullableDouble(double? val)
        {
            return val.HasValue ? "VAL: " + val : "NULL";
        }

        [ExcelFunction]
        public static string dnaNullableInt(int? val)
        {
            return val.HasValue ? "VAL: " + val : "NULL";
        }

        [ExcelFunction]
        public static string dnaNullableLong(long? val)
        {
            return val.HasValue ? "VAL: " + val : "NULL";
        }

        [ExcelFunction]
        public static string dnaNullableDateTime(DateTime? val)
        {
            return val.HasValue ? "VAL: " + val : "NULL";
        }

        [ExcelFunction]
        public static string dnaNullableBoolean(bool? val)
        {
            return val.HasValue ? "VAL: " + val : "NULL";
        }

        [ExcelFunction]
        public static string dnaNullableOptionalDateTime(DateTime? val = null)
        {
            return val.HasValue ? "VAL: " + val : "NULL";
        }

        public enum TestEnum1
        {
            Negative,
            Zero,
            Positive
        }

        public enum TestEnum2
        {
            Real,
            Imaginary
        }

        [ExcelFunction]
        public static TestEnum1 dnaReturnEnum1(string val)
        {
            //return val.HasValue ? val.Value : TestEnum.Zero;
            return (TestEnum1)Enum.Parse(typeof(TestEnum1), val, true);
        }

        [ExcelFunction]
        public static TestEnum2 dnaReturnEnum2(string val)
        {
            //return val.HasValue ? val.Value : TestEnum.Zero;
            return (TestEnum2)Enum.Parse(typeof(TestEnum2), val, true);
        }

        [ExcelFunction]
        public static Complex dnaEnumParameters(TestEnum1 val1, TestEnum2 val2)
        {
            //return val.HasValue ? val.Value : TestEnum.Zero;
            double r = 0;
            if (val1 == TestEnum1.Negative)
                r = -1;
            else if (val1 == TestEnum1.Positive)
                r = +1;
            double c = 0;
            if (val2 == TestEnum2.Imaginary)
                c = 1;
            return new Complex(r, c);
        }

        [ExcelFunction]
        public static Complex dnaNullableEnum(TestEnum1? val1, TestEnum2? val2)
        {
            //return val.HasValue ? val.Value : TestEnum.Zero;
            double r = 0;
            if (val1.HasValue && val1 == TestEnum1.Negative)
                r = -1;
            else if (val1 == TestEnum1.Positive)
                r = +1;
            double c = 0;
            if (val2.HasValue && val2 == TestEnum2.Imaginary)
                c = 1;
            return new Complex(r, c);
        }

        [ExcelFunction]
        public static Complex dnaComplex(Complex c)
        {
            //return val.HasValue ? val.Value : TestEnum.Zero;
            return c;
        }

        [ExcelFunction]
        public static Complex dnaNullableComplex(Complex? c)
        {
            //return val.HasValue ? val.Value : TestEnum.Zero;
            return c ?? new Complex(111, 222);
        }

        [MapArrayFunction]
        public static IEnumerable<TestEnum1> dnaEnumsEnumerated(IEnumerable<TestEnum2> v)
        {
            foreach (var i in v)
            {
                switch (i)
                {
                    case TestEnum2.Imaginary:
                        yield return TestEnum1.Negative;
                        break;
                    case TestEnum2.Real:
                        yield return TestEnum1.Positive;
                        break;
                }
            }
        }
    }

    // Here I test some custom conversions, including a two-hop conversion
    public class TestType1
    {
        public string Value;
        public TestType1(string value)
        {
            Value = value;
        }

        public override string ToString()
        {
            return "From Type 1 with " + Value;
        }

        [ExcelFunction]
        public static string dnaTestFunction1(TestType1 tt)
        {
            return "The Test (1) value is " + tt.Value;
        }
    }

    public class TestType2
    {
        readonly TestType1 _value;
        public TestType2(TestType1 value)
        {
            _value = value;
        }

        // Must be converted using TestType1 => TestType2, then string => TestType1
        [ExcelFunction]
        public static string dnaTestFunction2(TestType2 tt)
        {
            return "The Test (2) value is " + tt._value.Value;
        }

        [ExcelFunction]
        public static TestType1 dnaTestFunction2Ret1(TestType2 tt)
        {
            return new TestType1("The Test (2) value is " + tt._value.Value);
        }

        [ExcelFunction]
        public static string dnaJoinStrings(string separator, string[] values)
        {
            return string.Join(separator, values);
        }
    }

    public class ReturnTests
    {
        //[ExcelFunction]
        //public static object GetErrorNA()
        //{
        //    return ExcelError.ExcelErrorNA;
        //}

        [ExcelFunction]
        public static System.Threading.Tasks.Task<string> GetErrorNA(string s1)
        {
            return System.Threading.Tasks.Task.FromResult(s1);
        }
    }
}
