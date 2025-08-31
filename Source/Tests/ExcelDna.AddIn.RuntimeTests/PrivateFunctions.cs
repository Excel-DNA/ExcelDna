using ExcelDna.Integration;

namespace ExcelDna.AddIn.RuntimeTests
{
    public class ParameterClass { }

    public class PrivateFunctions
    {
        [ExcelFunction]
        internal static string InternalFunction(ParameterClass c)
        {
            return "";
        }

        [ExcelFunction]
        public string InstanceFunction(ParameterClass c)
        {
            return "";
        }
    }

    internal class InternalClass2
    {
        [ExcelFunction]
        public static string InternalClass(ParameterClass c)
        {
            return "";
        }
    }
}
