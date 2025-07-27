using ExcelDna.Registration;

namespace ExcelDna.AddIn.RuntimeTests
{
    internal class DynamicFunctions
    {
        public static void Register()
        {
            ExcelFunctionRegistration[] functions = {
                CreateRegistration(nameof(DynamicSayHello)),
                CreateRegistration(nameof(DynamicOptionalDouble)),
            };

            ExcelRegistration.RegisterFunctions(ExcelRegistration.ProcessFunctions(functions));
        }

        private static string DynamicSayHello(string name)
        {
            return $"Dynamic Hello {name}";
        }

        private static string DynamicOptionalDouble(double d = 4.56)
        {
            return "Dynamic Optional VAL: " + d.ToString();
        }

        private static ExcelFunctionRegistration CreateRegistration(string name)
        {
            return new ExcelFunctionRegistration(typeof(DynamicFunctions).GetMethod(name, System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static));
        }
    }
}
