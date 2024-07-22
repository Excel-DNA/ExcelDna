using ExcelDna.Integration;

namespace ExcelDna.AddIn.RuntimeTests
{
    public class MyFunctions
    {
        [ExcelFunction]
        public static string SayHello(string name)
        {
            return $"Hello {name}";
        }

        [ExcelFunction, Logging(7)]
        public static string SayHelloWithLoggingID(string name)
        {
            return $"Hello {name}";
        }

        [ExcelFunction]
        public static string MyDouble(double d)
        {
            return d.ToString();
        }

        [ExcelFunction]
        public static string MyNullableDouble(double? d)
        {
            return "Nullable VAL: " + (d.HasValue ? d : "NULL");
        }

        [ExcelFunction]
        public static string MyOptionalDouble(double d = 1.23)
        {
            return "Optional VAL: " + d.ToString();
        }

        [ExcelFunction]
        public static string MyEnum(DateTimeKind e)
        {
            return "Enum VAL: " + e.ToString();
        }

        [ExcelFunction]
        public static DateTimeKind MyEnumReturn(string s)
        {
            return Enum.Parse<DateTimeKind>(s);
        }

        [MapArrayFunction]
        public static IEnumerable<string> MyMapArray(IEnumerable<DateTimeKind> a)
        {
            return a.Select(i => "Array element VAL: " + i.ToString());
        }

        [ExcelAsyncFunction]
        public static string MyAsyncHello(string name, int msToSleep)
        {
            Thread.Sleep(msToSleep);
            return $"Hello async {name}";
        }

        [ExcelAsyncFunction]
        public static async Task<string> MyAsyncTaskHello(string name, int msDelay)
        {
            await Task.Delay(msDelay);
            return $"Hello async task {name}";
        }

        [ExcelAsyncFunction]
        public static async Task<string> MyAsyncGettingData(string name, int msDelay)
        {
            await Task.Delay(msDelay);
            return $"Hello async task {name}";
        }

        [ExcelFunction]
        public static Task<string> MyTaskHello(string name)
        {
            return Task.FromResult($"Hello task {name}");
        }

        [ExcelFunction]
        public static string MyStringArray(string[] s)
        {
            return "StringArray VALS: " + string.Concat(s);
        }

        [ExcelFunction]
        public static string MyStringArray2D(string[,] s)
        {
            string result = "";
            for (int i = 0; i < s.GetLength(0); i++)
            {
                for (int j = 0; j < s.GetLength(1); j++)
                {
                    result += s[i, j];
                }

                result += " ";
            }

            return $"StringArray2D VALS: {result}";
        }

        [ExcelFunction]
        public static string MyTestType1(TestType1 tt)
        {
            return "The TestType1 value is " + tt.Value;
        }

        [ExcelFunction]
        public static string MyTestType2(TestType2 tt)
        {
            return "The TestType2 value is " + tt.Value;
        }

        [ExcelFunction]
        public static string MyVersion2(Version v)
        {
            return "The Version value with field count 2 is " + v.ToString(2);
        }

        [ExcelFunction]
        public static string MyFunctionExecutionLog()
        {
            return Logger.GetLog();
        }

        [ExcelFunction]
        public static Calc MyCreateCalc(double d1, double d2)
        {
            return new Calc(d1, d2);
        }

        [ExcelFunction]
        public static Calc MyCreateCalc2(double d1, double d2)
        {
            return new Calc(d1 * 2, d2 * 2);
        }

        [ExcelFunction]
        public static double MyCalcSum(Calc c)
        {
            return c.Sum();
        }

        [ExcelFunction]
        public static DisposableObject MyCreateDisposableObject(int x)
        {
            return new DisposableObject();
        }

        [ExcelFunction]
        public static int MyGetDisposableObjectsCount()
        {
            return DisposableObject.ObjectsCount;
        }

        [ExcelFunction(IsThreadSafe = true)]
        public static Calc MyCreateCalcTS(double d1, double d2)
        {
            Thread.Sleep((int)d1);
            return new Calc(d1, d2);
        }

        [ExcelFunction(IsThreadSafe = true)]
        public static double MyCalcSumTS(Calc c)
        {
            Thread.Sleep((int)c.Sum());
            return c.Sum();
        }

        [ExcelFunction]
        public static string MyRange(Microsoft.Office.Interop.Excel.Range r)
        {
            return r.Address;
        }

        [ExcelFunction]
        public static string MyUserCodeConversionReferenceToRange([ExcelArgument(AllowReference = true)] object r)
        {
            return ExcelConversionUtil.ReferenceToRange((ExcelReference)r).Address;
        }

        [ExcelFunction]
        public static string MyParamsFunc1(
            [ExcelArgument(Name = "first.Input", Description = "is a useful start")]
            object input,
            [ExcelArgument(Description = "is another param start")]
            string QtherInpEt,
            [ExcelArgument(Name = "Value", Description = "gives the Rest")]
            params object[] args)
        {
            return input + "," + QtherInpEt + ", : " + args.Length;
        }

        [ExcelFunction]
        public static string MyParamsFunc2(
            [ExcelArgument(Name = "first.Input", Description = "is a useful start")]
            object input,
            [ExcelArgument(Name = "second.Input", Description = "is some more stuff")]
            string input2,
            [ExcelArgument(Description = "is another param ")]
            string QtherInpEt,
            [ExcelArgument(Name = "Value", Description = "gives the Rest")]
            params object[] args)
        {
            var content = string.Join(",", args.Select(ValueType => ValueType.ToString()));
            return input + "," + input2 + "," + QtherInpEt + ", " + $"[{args.Length}: {content}]";
        }

        [ExcelFunction]
        public static string MyParamsJoinString(string separator, params string[] values)
        {
            return String.Join(separator, values);
        }
    }
}
