using System.Runtime.InteropServices;
using ExcelDna.ComInterop;
using ExcelDna.Integration;

namespace NET4Dscom
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IComLibrary
    {
        string ComLibraryHello();
        double Add(double x, double y);
    }

    [ComDefaultInterface(typeof(IComLibrary))]
    public class ComLibrary : IComLibrary
    {
        public string ComLibraryHello()
        {
            return "Hello from NET4Dscom.ComLibrary at " + ExcelDnaUtil.XllPath;
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
        public static object NET4DscomHello()
        {
            return "Hello from NET4Dscom!";
        }
    }
}
