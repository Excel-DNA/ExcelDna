using ExcelDna.Integration;
using ExcelDna.Registration;

[assembly: ExcelHandleExternal(typeof(System.Reflection.Assembly))]

namespace ExcelDna.AddIn.RuntimeTests
{
    public class MyFunctions
    {
        [ExcelCommand(MenuText = "MyCommandHello")]
        public static void MyCommandHello()
        {
            Logger.Log("Hello command.");
        }

        [ExcelFunction]
        public static string SayHello(string name)
        {
            return $"Hello {name}";
        }

        [ExcelFunction]
        public static string MySayHelloWithExclamation(string name)
        {
            return $"Hello with exclamation {name}";
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
        public static string MyDateTime(DateTime d)
        {
            return d.ToString();
        }

        [ExcelFunction]
        public static string MyNullableDouble(double? d)
        {
            return "Nullable VAL: " + (d.HasValue ? d : "NULL");
        }

        [ExcelFunction]
        public static string MyNullableDateTime(DateTime? dt)
        {
            return "Nullable DateTime: " + (dt.HasValue ? dt : "NULL");
        }

        [ExcelFunction]
        public static string MyOptionalDouble(double d = 1.23)
        {
            return "Optional VAL: " + d.ToString();
        }

        [ExcelFunction]
        public static string MyOptionalDateTime(DateTime dt = default)
        {
            return "Optional DateTime: " + dt.ToString();
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
        public static TestType1 MyReturnTestType1(string s)
        {
            return new TestType1("The TestType1 return value is " + s);
        }

        [ExcelFunction]
        public static string MyFunctionExecutionLog()
        {
            string result = Logger.GetLog();
            Logger.ClearLog();
            return result;
        }

        [ExcelFunction]
        [return: ExcelHandle]
        public static Calc MyCreateCalc(double d1, double d2)
        {
            return new Calc(d1, d2);
        }

        [ExcelFunction]
        [return: ExcelHandle]
        public static Calc MyCreateCalc2(double d1, double d2)
        {
            return new Calc(d1 * 2, d2 * 2);
        }

        [ExcelFunction]
        public static CalcExcelHandle MyCreateCalcExcelHandle(double d1, double d2)
        {
            return new CalcExcelHandle(d1, d2);
        }

        [ExcelFunction]
        public static CalcStructExcelHandle MyCreateCalcStructExcelHandle(double d1, double d2)
        {
            return new CalcStructExcelHandle(d1, d2);
        }

        [ExcelFunction]
        public static CalcExcelHandleExternal MyCreateCalcExcelHandleExternal(double d1, double d2)
        {
            return new CalcExcelHandleExternal(d1, d2);
        }

        [ExcelFunction]
        public static System.Reflection.Assembly MyGetExecutingAssembly()
        {
            return System.Reflection.Assembly.GetExecutingAssembly();
        }

        [ExcelFunction]
        [return: ExcelHandle]
        public static int MyCreateSquareIntObject(int i)
        {
            return i * i;
        }

        [ExcelFunction]
        [return: ExcelHandle]
        public static async Task<Calc> MyTaskCreateCalc(int millisecondsDelay, double d1, double d2)
        {
            await Task.Delay(millisecondsDelay);
            return new Calc(d1, d2);
        }

        [ExcelFunction]
        public static async Task<CalcExcelHandle> MyTaskCreateCalcExcelHandle(int millisecondsDelay, double d1, double d2)
        {
            await Task.Delay(millisecondsDelay);
            return new CalcExcelHandle(d1, d2);
        }

        [ExcelFunction]
        [return: ExcelHandle]
        public static async Task<Calc> MyTaskCreateCalcWithCancellation(int millisecondsDelay, double d1, double d2, CancellationToken ct)
        {
            await Task.Delay(millisecondsDelay);
            return new Calc(d1, d2);
        }

        [ExcelAsyncFunction]
        [return: ExcelHandle]
        public static Calc MyAsyncCreateCalc(int millisecondsDelay, double d1, double d2)
        {
            if (millisecondsDelay > 0)
                Thread.Sleep(millisecondsDelay);

            return new Calc(d1, d2);
        }

        [ExcelAsyncFunction]
        [return: ExcelHandle]
        public static Calc MyAsyncCreateCalcWithCancellation(int millisecondsDelay, double d1, double d2, CancellationToken ct)
        {
            if (millisecondsDelay > 0)
                Thread.Sleep(millisecondsDelay);

            return new Calc(d1, d2);
        }

        [ExcelFunction]
        public static double MyCalcSum([ExcelHandle] Calc c)
        {
            return c.Sum();
        }

        [ExcelFunction]
        public static double MyCalcExcelHandleMul(CalcExcelHandle c)
        {
            return c.Mul();
        }

        [ExcelFunction]
        public static double MyCalcStructExcelHandleMul(CalcStructExcelHandle c)
        {
            return c.Mul();
        }

        [ExcelFunction]
        public static double MyCalcExcelHandleExternalMul(CalcExcelHandleExternal c)
        {
            return c.Mul();
        }

        [ExcelFunction]
        public static string? MyGetAssemblyName(System.Reflection.Assembly assembly)
        {
            return assembly.GetName().Name;
        }

        [ExcelFunction]
        public static Task<double> MyTaskCalcSum([ExcelHandle] Calc c)
        {
            return Task.FromResult(c.Sum());
        }

        [ExcelFunction]
        public static Task<double> MyTaskCalcDoubleSumWithCancellation([ExcelHandle] Calc c, CancellationToken ct)
        {
            return Task.FromResult(c.Sum() * 2);
        }

        [ExcelFunction]
        public static string MyPrintIntObject([ExcelHandle] int i)
        {
            return $"IntObject value={i}";
        }

        [ExcelFunction]
        public static string MyPrintMixedIntObject(double d, [ExcelHandle] int i)
        {
            return $"double value={d}, IntObject value={i}";
        }

        [ExcelFunction]
        public static IObservable<string> MyStringObservable(string s)
        {
            return new ObservableString(s);
        }

        [ExcelFunction]
        [return: ExcelHandle]
        public static IObservable<Calc> MyCalcObservable(double d1, double d2)
        {
            return new ObservableCalc(d1, d2);
        }

        [ExcelFunction]
        public static IObservable<CalcExcelHandle> MyCalcExcelHandleObservable(double d1, double d2)
        {
            return new ObservableCalcExcelHandle(d1, d2);
        }

        [ExcelFunction]
        public static IObservable<string> MyCalcSumObservable([ExcelHandle] Calc c)
        {
            return new ObservableString(c.Sum().ToString());
        }

        [ExcelFunction]
        [return: ExcelHandle]
        public static DisposableObject MyCreateDisposableObject(int x)
        {
            return new DisposableObject();
        }

        [ExcelFunction]
        public static int MyGetDisposableObjectsCount()
        {
            return DisposableObject.ObjectsCount;
        }

        [ExcelFunction]
        public static int MyGetCreatedDisposableObjectsCount()
        {
            return DisposableObject.CreatedObjectsCount;
        }

        [ExcelFunction]
        [return: ExcelHandle]
        public static async Task<DisposableObject> MyTaskCreateDisposableObject(int millisecondsDelay, int x)
        {
            await Task.Delay(millisecondsDelay);
            return new DisposableObject();
        }

        [ExcelFunction(IsThreadSafe = true)]
        [return: ExcelHandle]
        public static Calc MyCreateCalcTS(double d1, double d2)
        {
            Thread.Sleep((int)d1);
            return new Calc(d1, d2);
        }

        [ExcelFunction(IsThreadSafe = true)]
        public static double MyCalcSumTS([ExcelHandle] Calc c)
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

        [ExcelFunction]
        public static string MyWindowHandle()
        {
            return $"My WindowHandle is {ExcelDnaUtil.WindowHandle}.";
        }

        [ExcelFunction]
        public static string MyEnum16(int i1, int i2, int i3, int i4, int i5, int i6, int i7, int i8, int i9, int i10, int i11, int i12, int i13, int i14, int i15, TestEnum e)
        {
            int sum = i1 + i2 + i3 + i4 + i5 + i6 + i7 + i8 + i9 + i10 + i11 + i12 + i13 + i14 + i15;
            return $"{e} {sum}";
        }

        [ExcelFunction]
        public static string MyEnum17(int i1, int i2, int i3, int i4, int i5, int i6, int i7, int i8, int i9, int i10, int i11, int i12, int i13, int i14, int i15, int i16, TestEnum e)
        {
            int sum = i1 + i2 + i3 + i4 + i5 + i6 + i7 + i8 + i9 + i10 + i11 + i12 + i13 + i14 + i15 + i16;
            return $"{e} {sum}";
        }
    }
}
