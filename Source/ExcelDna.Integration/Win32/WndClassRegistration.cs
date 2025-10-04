#if !USE_WINDOWS_FORMS

using System.Collections.Generic;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.UI.WindowsAndMessaging;

namespace ExcelDna.Integration.Win32
{
    internal class WndClassRegistration
    {
        public WndClassRegistration(string name)
        {
            this.Name = name;

            unsafe
            {
                fixed (char* pClassName = name)
                {
                    wndClass = new();
                    wndClass.lpfnWndProc = ClassWndProc;
                    wndClass.hInstance = (HINSTANCE)ExcelDnaUtil.ModuleXll;
                    wndClass.lpszClassName = new PCWSTR(pClassName);
                    PInvoke.RegisterClass(wndClass);
                }
            }
        }

        public string Name { get; }

        public void RegisterWnd(HWND hWND, WNDPROC wndProc)
        {
            registeredWndProc[hWND] = wndProc;
        }

        public void UnRegisterWnd(HWND hWND)
        {
            registeredWndProc.Remove(hWND);
        }

        private LRESULT ClassWndProc(HWND hWnd, uint msg, WPARAM wParam, LPARAM lParam)
        {
            if (registeredWndProc.TryGetValue(hWnd, out WNDPROC wndProc))
                wndProc(hWnd, msg, wParam, lParam);

            return PInvoke.DefWindowProc(hWnd, msg, wParam, lParam);
        }

        private WNDCLASSW wndClass;
        private Dictionary<HWND, WNDPROC> registeredWndProc = new();
    }
}

#endif
