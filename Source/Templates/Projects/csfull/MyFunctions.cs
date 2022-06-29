using ExcelDna.Integration;

namespace csfull;

public static class MyFunctions
{
    [ExcelFunction(Description = "My first .NET function")]
    public static string SayHello(string name)
    {
        return "Hello " + name;
    }
}
