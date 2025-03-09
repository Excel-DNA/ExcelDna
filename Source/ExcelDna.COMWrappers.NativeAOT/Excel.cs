using Addin.ComApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDna.COMWrappers.NativeAOT
{
    internal class Excel
    {
        public static object? GetApplicationFromWindow(IntPtr hWndMain)
        {
            object? result = null;
            StringBuilder cname = new StringBuilder(256);

            EnumChildWindows(hWndMain, delegate (IntPtr hWndEnum, IntPtr param)
            {
                // Check the window class
                GetClassNameW(hWndEnum, cname, cname.Capacity);

                System.Diagnostics.Trace.WriteLine("[GetApplicationFromWindow cname ]" + cname.ToString());


                if (cname.ToString() != "EXCEL7")
                    // Not a workbook window, continue enumerating
                    return true;

                System.Diagnostics.Trace.WriteLine("[GetApplicationFromWindow got EXCEL7]");


                IntPtr pUnk = IntPtr.Zero;
                int hr = AccessibleObjectFromWindow(hWndEnum, OBJID_NATIVEOM, IID_IDispatchBytes, ref pUnk);
                if (hr != 0)
                {
                    // Window does not implement the IID, continue enumerating
                    return true;
                }

                System.Diagnostics.Trace.WriteLine("[GetApplicationFromWindow got pUnk]");

                // Marshal to .NET, then call .Application
                //object obj = Marshal.GetObjectForIUnknown(pUnk);
                //Console.WriteLine("GetApplicationFromWindow GetObjectForIUnknown");

                ComWrappers cw = new StrategyBasedComWrappers();
                //object obj = cw.GetOrCreateObjectForComInstance(pUnk, CreateObjectFlags.None);
                //ComObject co = obj;

                //var t = obj.GetType();
                IDispatch foo = (IDispatch)cw.GetOrCreateObjectForComInstance(pUnk, CreateObjectFlags.None);
                //foo.
                //IntPtr app = foo.get_Application2xx();

                //hr = foo.GetTypeInfoCount(out uint tcount);

                //const int LcidUsEnglish = 0x0409;
                //string[] names = new string[1];
                //int[] ids = new int[1];
                //names[0] = "Application";
                //Guid g = Guid.Empty;
                //hr = foo.GetIDsOfNames(ref g, names, 1, LcidUsEnglish, ids);

                //int[] i1 = new int[4];
                //IntPtr i2 = 0;
                //hr = foo.Invoke(ids[0], ref g, LcidUsEnglish, 2, i1, out IntPtr pVarResult, ref i2, out uint perr);

                var excelWindowWrapper = new ExcelObject(foo);

                var app = excelWindowWrapper.GetProperty("Application");
                if (app is ExcelObject excelObject)
                    result = excelObject._interfacePtr;
                else
                    result = app;

                //dynamic dapp = app;
                //var commandBars = dapp.CommandBars;
                //var worksheetBar = commandBars[1];
                //var controls = worksheetBar.Controls;

                //string menuName = "ConsoleApp1 Menu";
                //var menu = controls.AddPopup(menuName);
                //menu.Caption = menuName;

                //Console.WriteLine("GetApplicationFromWindow GetOrCreateObjectForComInstance");


                return result == null;
            }, IntPtr.Zero);

            System.Diagnostics.Trace.WriteLine("[GetApplicationFromWindow result]" + result?.GetType().ToString());

            return result;
        }

        private delegate bool EnumWindowsCallback(IntPtr hwnd, /*ref*/ IntPtr param);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsCallback callback, /*ref*/ IntPtr param);

        [DllImport("user32.dll")]
        private static extern int GetClassNameW(IntPtr hwnd, [MarshalAs(UnmanagedType.LPWStr)] StringBuilder buf, int nMaxCount);

        [DllImport("Oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(
      IntPtr hwnd, uint dwObjectID, byte[] riid,
      ref IntPtr ptr /*ppUnk*/);

        private const uint OBJID_NATIVEOM = 0xFFFFFFF0;

        private static readonly byte[] IID_IDispatchBytes = new Guid("{00020400-0000-0000-C000-000000000046}").ToByteArray();
    }
}
