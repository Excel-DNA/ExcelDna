using ExcelDna.Integration;

namespace NET6CustomRuntimeConfiguration
{
    public static class Functions
    {
        [ExcelFunction]
        public static string SayHello(string name)
        {
            var webApplication = WebApplication.Create();
            return $"{name} WebApplication Environment: {webApplication.Environment.EnvironmentName}.";
        }
    }
}
