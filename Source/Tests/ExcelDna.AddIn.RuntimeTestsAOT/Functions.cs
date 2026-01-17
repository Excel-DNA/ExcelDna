using ExcelDna.Integration;
using ExcelDna.Registration;

[assembly: ExcelHandleExternal(typeof(System.Reflection.Assembly))]

namespace ExcelDna.AddIn.RuntimeTestsAOT
{
    public class Functions
    {
        [ExcelCommand(MenuText = "NativeCommandHello")]
        public static void NativeCommandHello()
        {
            Logger.Log("Native hello command.");
        }

        [ExcelFunction]
        public static string NativeHello(string name)
        {
            return $"Hello {name}!";
        }

        [ExcelFunction]
        public static int NativeSum(int i1, int i2)
        {
            return i1 + i2;
        }

        [ExcelAsyncFunction]
        public static async Task<string> NativeAsyncTaskHello(string name, int msDelay)
        {
            await Task.Delay(msDelay);
            return $"Hello native async task {name}";
        }

        [ExcelAsyncFunction]
        public static string NativeAsyncHello(string name, int msToSleep)
        {
            Thread.Sleep(msToSleep);
            return $"Hello native async {name}";
        }

        [ExcelFunction]
        public static Task<string> NativeTaskHello(string name)
        {
            return Task.FromResult($"Hello native task {name}");
        }

        [ExcelFunction]
        public static Task<bool> NativeTaskBool()
        {
            return Task.FromResult(true);
        }

        [ExcelFunction]
        public static Task<CalcExcelHandle> NativeTaskCalcExcelHandle(double d1, double d2)
        {
            return Task.FromResult(new CalcExcelHandle(d1, d2));
        }

        [ExcelFunction]
        public static Task<CalcExcelHandle> NativeTaskCalcExcelHandleWithCancellation(double d1, double d2, CancellationToken cancellation)
        {
            return Task.FromResult(new CalcExcelHandle(d1, d2));
        }

        [ExcelFunction]
        public static Task<bool> NativeTaskBoolWithCancellation(CancellationToken cancellation)
        {
            return Task.FromResult(true);
        }

        [ExcelAsyncFunction]
        public static bool NativeAsyncBool()
        {
            return true;
        }

        [ExcelAsyncFunction]
        public static bool NativeAsyncBoolWithCancellation(CancellationToken cancellation)
        {
            return true;
        }

        [ExcelAsyncFunction]
        public static CalcExcelHandle NativeAsyncCalcExcelHandle(double d1, double d2)
        {
            return new CalcExcelHandle(d1, d2);
        }

        [ExcelAsyncFunction]
        public static CalcExcelHandle NativeAsyncCalcExcelHandleWithCancellation(double d1, double d2, CancellationToken cancellation)
        {
            return new CalcExcelHandle(d1, d2);
        }

        [ExcelFunction]
        public static string NativeApplicationName()
        {
            return (string)ExcelDnaUtil.DynamicApplication.Get("Name")!;
        }

        [ExcelFunction]
        public static double NativeApplicationGetCellValue(string cell)
        {
            var workbook = (IDynamic)ExcelDnaUtil.DynamicApplication.Get("ActiveWorkbook")!;
            var sheets = (IDynamic)workbook.Get("Sheets")!;
            var sheet = (IDynamic)sheets[1]!;
            var range = (IDynamic)sheet.Get("Range", [cell])!;
            return (double)range.Get("Value")!;
        }

        [ExcelFunction]
        public static double NativeApplicationGetCellValueT(string cell)
        {
            var workbook = ExcelDnaUtil.DynamicApplication.Get<IDynamic>("ActiveWorkbook");
            var sheets = workbook.Get<IDynamic>("Sheets");
            var sheet = (IDynamic)sheets[1]!;
            var range = sheet.Get<IDynamic>("Range", [cell]);
            return range.Get<double>("Value");
        }

        [ExcelFunction]
        public static int NativeApplicationAlignCellRight(string cell)
        {
            var workbook = ExcelDnaUtil.DynamicApplication.Get<IDynamic>("ActiveWorkbook");
            var sheets = workbook.Get<IDynamic>("Sheets");
            var sheet = (IDynamic)sheets[1]!;
            var range = sheet.Get<IDynamic>("Range", [cell]);
            range.Set("HorizontalAlignment", -4152);
            return range.Get<int>("HorizontalAlignment");
        }

        [ExcelFunction]
        public static string NativeApplicationAddCellComment(string cell, string comment)
        {
            var workbook = ExcelDnaUtil.DynamicApplication.Get<IDynamic>("ActiveWorkbook");
            var sheets = workbook.Get<IDynamic>("Sheets");
            var sheet = (IDynamic)sheets[1]!;
            var range = sheet.Get<IDynamic>("Range", [cell]);
            var newComment = (IDynamic)range.Invoke("AddComment", [comment])!;
            return newComment.Invoke<string>("Text", []);
        }

        [ExcelFunction]
        public static object NativeRangeConcat2(object[,] values)
        {
            string result = "";
            int rows = values.GetLength(0);
            int cols = values.GetLength(1);
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    object value = values[i, j];
                    result += value.ToString();
                }
            }
            return result;
        }

        [ExcelFunction]
        public static string NativeNullableDouble(double? d)
        {
            return "Native Nullable VAL: " + (d.HasValue ? d : "NULL");
        }

        [ExcelFunction]
        public static string NativeOptionalDouble(double d = 1.23)
        {
            return "Native Optional VAL: " + d.ToString();
        }

        [ExcelFunction]
        public static string NativeRangeAddress(IRange r)
        {
            return "Native Address: " + r.Get<string>("Address");
        }

        [ExcelFunction]
        public static string NativeEnum(DateTimeKind e)
        {
            return "Native Enum VAL: " + e.ToString();
        }

        [ExcelFunction]
        public static DateTimeKind NativeEnumReturn(string s)
        {
            return Enum.Parse<DateTimeKind>(s);
        }

        [ExcelFunction]
        public static string NativeStringArray(string[] s)
        {
            return "Native StringArray VALS: " + string.Concat(s);
        }

        [ExcelFunction]
        public static string NativeStringArray2D(string[,] s)
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

            return $"Native StringArray2D VALS: {result}";
        }

        [ExcelFunction]
        public static string NativeParamsFunc1(
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
        public static string NativeParamsFunc2(
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
        public static string NativeParamsJoinString(string separator, params string[] values)
        {
            return String.Join(separator, values);
        }

        [ExcelFunction]
        [return: ExcelHandle]
        public static Calc NativeCreateCalc(double d1, double d2)
        {
            return new Calc(d1, d2);
        }

        [ExcelFunction]
        public static double NativeCalcSum([ExcelHandle] Calc c)
        {
            return c.Sum();
        }

        [ExcelFunction]
        public static CalcExcelHandle NativeCreateCalcExcelHandle(double d1, double d2)
        {
            return new CalcExcelHandle(d1, d2);
        }

        [ExcelFunction]
        public static double NativeCalcExcelHandleMul(CalcExcelHandle c)
        {
            return c.Mul();
        }

        [ExcelFunction]
        public static System.Reflection.Assembly NativeGetExecutingAssembly()
        {
            return System.Reflection.Assembly.GetExecutingAssembly();
        }

        [ExcelFunction]
        public static string? NativeGetAssemblyName(System.Reflection.Assembly assembly)
        {
            return assembly.GetName().Name;
        }

        [ExcelFunction]
        public static string NativeVersion2(Version v)
        {
            return "The Native Version value with field count 2 is " + v.ToString(2);
        }

        [ExcelFunction]
        public static TestType1 NativeReturnTestType1(string s)
        {
            return new TestType1("The Native TestType1 return value is " + s);
        }

        [ExcelFunction]
        public static string NativeFunctionExecutionLog()
        {
            string result = Logger.GetLog();
            Logger.ClearLog();
            return result;
        }

        [ExcelFunction, Logging(7)]
        public static string NativeSayHelloWithLoggingID(string name)
        {
            return $"Native Hello {name}";
        }

        [ExcelFunction]
        public static string NativeWindowHandle()
        {
            return $"Native WindowHandle is {ExcelDnaUtil.WindowHandle}.";
        }

        [ExcelFunction]
        public static IObservable<string> NativeStringObservable(string s)
        {
            return new ObservableString(s);
        }

        [ExcelFunction]
        [return: ExcelHandle]
        public static IObservable<Calc> NativeCalcObservable(double d1, double d2)
        {
            return new ObservableCalc(d1, d2);
        }

        [ExcelFunction]
        public static IObservable<string> NativeCalcSumObservable([ExcelHandle] Calc c)
        {
            return new ObservableString(c.Sum().ToString());
        }
    }
}
