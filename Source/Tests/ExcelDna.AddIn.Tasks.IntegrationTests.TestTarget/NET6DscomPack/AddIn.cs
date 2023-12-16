using System.Runtime.InteropServices;
using ExcelDna.ComInterop;
using ExcelDna.Integration;

namespace NET6DscomPack
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IComLibrary
    {
        string ComLibraryHello();
    }

    [ComDefaultInterface(typeof(IComLibrary))]
    public class ComLibrary : IComLibrary
    {
        public string ComLibraryHello()
        {
            return "Hello from NET6DscomPack.ComLibrary at " + ExcelDnaUtil.XllPath;
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
}
