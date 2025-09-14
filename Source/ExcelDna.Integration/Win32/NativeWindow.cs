#if !USE_WINDOWS_FORMS

using System.Collections.Generic;
using System.Runtime.InteropServices;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.UI.WindowsAndMessaging;

namespace ExcelDna.Integration.Win32
{
    internal abstract class NativeWindow
    {
        static NativeWindow()
        {
            unsafe
            {
                fixed (char* pClassName = className)
                {
                    wndClass = new();
                    wndClass.lpfnWndProc = ClassWndProc;
                    wndClass.hInstance = (HINSTANCE)ExcelDnaUtil.ModuleXll;
                    wndClass.lpszClassName = new PCWSTR(pClassName);
                    PInvoke.RegisterClass(wndClass);
                }
            }
        }

        public System.IntPtr Handle { get; private set; }

        public virtual void CreateHandle(CreateParams cp)
        {
            unsafe
            {
                HWND hwnd = PInvoke.CreateWindowEx(0, className, null, 0, 0, 0, 1, 1, new HWND(cp.Parent), null, new InstanceHandle(ExcelDnaUtil.ModuleXll), null);
                if (hwnd == 0)
                    throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());

                Handle = hwnd;
                registeredWndProc[hwnd] = InstanceWndProc;
            }
        }

        public virtual void DestroyHandle()
        {
            HWND hwnd = new HWND(Handle);
            PInvoke.DestroyWindow(hwnd);
            registeredWndProc.Remove(hwnd);
            Handle = 0;
        }

        protected virtual void WndProc(ref Message m)
        {
        }

        private LRESULT InstanceWndProc(HWND hWnd, uint msg, WPARAM wParam, LPARAM lParam)
        {
            Message message = new Message() { Msg = (int)msg };
            WndProc(ref message);
            return new LRESULT();
        }

        private static LRESULT ClassWndProc(HWND hWnd, uint msg, WPARAM wParam, LPARAM lParam)
        {
            if (registeredWndProc.TryGetValue(hWnd, out WNDPROC wndProc))
                wndProc(hWnd, msg, wParam, lParam);

            return PInvoke.DefWindowProc(hWnd, msg, wParam, lParam);
        }

        const string className = "ExcelDna.Integration.Win32.NativeWindow";
        private static WNDCLASSW wndClass;
        private static Dictionary<HWND, WNDPROC> registeredWndProc = new();
    }
}

#endif
