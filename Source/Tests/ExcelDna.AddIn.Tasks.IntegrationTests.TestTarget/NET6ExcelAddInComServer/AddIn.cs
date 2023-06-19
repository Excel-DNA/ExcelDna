using System.Runtime.InteropServices;
using ExcelDna.ComInterop;
using ExcelDna.Integration;

namespace NET6ExcelAddInComServer
{
    public interface IComLibraryInterface
    {
        public string ComLibraryHello();
        public double Add(double x, double y);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class ComLibrary : IComLibraryInterface
    {
        public string ComLibraryHello()
        {
            return "Hello from NET6 ExcelAddInComServer.ComLibrary";
        }

        public double Add(double x, double y)
        {
            return x + y;
        }
    }

    [ComVisible(false)]
    public class ExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ComServer.DllRegisterServer();
        }
        public void AutoClose()
        {
            ComServer.DllUnregisterServer();
        }
    }

    public static class Functions
    {
        [ExcelFunction]
        public static object ExcelAddInComServerHello()
        {
            return "Hello from NET6 ExcelAddInComServer!";
        }
    }
}
