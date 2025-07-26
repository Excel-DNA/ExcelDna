using ExcelDna.Registration;

namespace ExcelDna.AddIn.RuntimeTests
{
    internal class DynamicFunctions
    {
        public static void Register()
        {
            ExcelFunctionRegistration[] functions = { CreateRegistration(nameof(DynamicSayHello)) };
            ExcelRegistration.RegisterFunctions(functions);
        }

        private static string DynamicSayHello(string name)
        {
            return $"Dynamic Hello {name}";
        }

        private static ExcelFunctionRegistration CreateRegistration(string name)
        {
            return new ExcelFunctionRegistration(typeof(DynamicFunctions).GetMethod(name, System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static));
        }
    }
}
