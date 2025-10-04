#if !USE_WINDOWS_FORMS

using System.Runtime.InteropServices;
using Windows.Win32;
using Windows.Win32.Foundation;

namespace ExcelDna.Integration.Win32
{
    internal abstract class NativeWindow
    {
        public System.IntPtr Handle { get; private set; }

        public virtual void CreateHandle(CreateParams cp)
        {
            unsafe
            {
                HWND hwnd = PInvoke.CreateWindowEx(0, wndClassRegistration.Name, null, 0, 0, 0, 1, 1, new HWND(cp.Parent), null, new InstanceHandle(ExcelDnaUtil.ModuleXll), null);
                if (hwnd == 0)
                    throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());

                Handle = hwnd;
                wndClassRegistration.RegisterWnd(hwnd, InstanceWndProc);
            }
        }

        public virtual void DestroyHandle()
        {
            HWND hwnd = new HWND(Handle);
            PInvoke.DestroyWindow(hwnd);
            wndClassRegistration.UnRegisterWnd(hwnd);
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

        private static WndClassRegistration wndClassRegistration = new("ExcelDna.Integration.Win32.NativeWindow");
    }
}

#endif
