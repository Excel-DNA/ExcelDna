#if !USE_WINDOWS_FORMS

using System.Collections.Generic;
using System.Runtime.InteropServices;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.UI.WindowsAndMessaging;
using static ExcelDna.Integration.Win32.Constants;

namespace ExcelDna.Integration.Win32
{
    internal class LogDisplayWindow
    {
        static LogDisplayWindow()
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

        public LogDisplayWindow(CreateParams cp, string text)
        {
            unsafe
            {
                string name = DnaLibrary.CurrentLibraryName + " - Diagnostic Display";
                wnd = PInvoke.CreateWindowEx(0, className, name, WINDOW_STYLE.WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, new HWND(cp.Parent), null, new InstanceHandle(ExcelDnaUtil.ModuleXll), null);
                if (wnd == 0)
                    throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());

                registeredWndProc[wnd] = InstanceWndProc;

                editWnd = PInvoke.CreateWindowEx(0, "EDIT", text, WINDOW_STYLE.WS_CHILD | WINDOW_STYLE.WS_VISIBLE | WINDOW_STYLE.WS_VSCROLL | WINDOW_STYLE.WS_HSCROLL | (WINDOW_STYLE)ES_MULTILINE, 0, 0, 0, 0, wnd, null, new InstanceHandle(ExcelDnaUtil.ModuleXll), null);
            }
        }

        public void Show()
        {
            PInvoke.ShowWindow(wnd, SHOW_WINDOW_CMD.SW_SHOWNORMAL);
        }

        public void SetText(string text)
        {
            unsafe
            {
                fixed (char* pText = text)
                {
                    PInvoke.SetWindowText(editWnd, new PCWSTR(pText));
                }
            }
        }

        private LRESULT InstanceWndProc(HWND hWnd, uint msg, WPARAM wParam, LPARAM lParam)
        {
            switch (msg)
            {
                case WM_SIZE:
                    PInvoke.MoveWindow(editWnd, 0, 0, Util.LoWord((int)lParam), Util.HiWord((int)lParam), true);
                    break;
            }

            return new LRESULT();
        }

        private static LRESULT ClassWndProc(HWND hWnd, uint msg, WPARAM wParam, LPARAM lParam)
        {
            if (registeredWndProc.TryGetValue(hWnd, out WNDPROC wndProc))
                wndProc(hWnd, msg, wParam, lParam);

            return PInvoke.DefWindowProc(hWnd, msg, wParam, lParam);
        }

        const string className = "ExcelDna.Integration.Win32.LogDisplayWindow";
        private static WNDCLASSW wndClass;
        private static Dictionary<HWND, WNDPROC> registeredWndProc = new();

        private HWND wnd;
        private HWND editWnd;
    }
}

#endif
