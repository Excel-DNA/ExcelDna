/*
  Copyright (C) 2005-2008 Govert van Drimmelen

  This software is provided 'as-is', without any express or implied
  warranty.  In no event will the authors be held liable for any damages
  arising from the use of this software.

  Permission is granted to anyone to use this software for any purpose,
  including commercial applications, and to alter it and redistribute it
  freely, subject to the following restrictions:

  1. The origin of this software must not be misrepresented; you must not
     claim that you wrote the original software. If you use this software
     in a product, an acknowledgment in the product documentation would be
     appreciated but is not required.
  2. Altered source versions must be plainly marked as such, and must not be
     misrepresented as being the original software.
  3. This notice may not be removed or altered from any source distribution.


  Govert van Drimmelen
  govert@icon.co.za
*/

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelDna.Integration
{
    [Obsolete("Use ExcelDna.Integration.ExcelDnaUtil class.")]
    public class Excel
    {
        [Obsolete("Use ExcelDna.Integration.ExcelDnaUtil.WindowHandle property.")]
        public static IntPtr WindowHandle
        {
            get { return ExcelDnaUtil.WindowHandle; }
        }

        [Obsolete("Use ExcelDna.Integration.ExcelDnaUtil.Application property.")]
        public static object Application
        {
            get { return ExcelDnaUtil.Application; }
        }

        [Obsolete("Use ExcelDna.Integration.ExcelDnaUtil.IsInFunctionWizard property.")]
        public static bool IsInFunctionWizard()
        {
            return ExcelDnaUtil.IsInFunctionWizard();
        }
    }

	public class ExcelDnaUtil
	{
		private delegate bool EnumWindowsCallback(IntPtr hwnd, /*ref*/ IntPtr param);

		[DllImport("user32.dll")]
		private static extern int EnumWindows(EnumWindowsCallback callback, /*ref*/ IntPtr param);
		[DllImport("user32.dll")]
		private static extern IntPtr GetParent(IntPtr hwnd);
		[DllImport("user32.dll")]
		private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsCallback callback, /*ref*/ IntPtr param);
		[DllImport("user32.dll")]
		private static extern int GetClassName(IntPtr hwnd, [MarshalAs(UnmanagedType.LPStr)] StringBuilder buf, int nMaxCount);
		[DllImport("Oleacc.dll")]
		private static extern int AccessibleObjectFromWindow(
			  IntPtr hwnd, uint dwObjectID, byte[] riid,
			  ref IntPtr ptr /*ppUnk*/);

		private static IntPtr _hWndExcel = (IntPtr)0;
		public static IntPtr WindowHandle
		{
			get
			{
				// CONSIDER: Process.GetCurrentProcess().MainWindowHandle;
				if (_hWndExcel == (IntPtr)0)
				{
					// Get the LoWord
					short loWord = (short)XlCall.Excel(XlCall.xlGetHwnd);
					EnumWindows(delegate(IntPtr hWndEnum, IntPtr param)
						{
							// Check the loWord
							if (((uint)hWndEnum & 0x0000FFFF) == (uint)loWord)
							{
								// Check the window class
								StringBuilder cname = new StringBuilder(256);
								GetClassName(hWndEnum, cname, cname.Capacity);
								if (cname.ToString() == "XLMAIN")
								{
									_hWndExcel = hWndEnum;
									return false;	// Stop enumerating
								}
							}
							return true;	// Continue enumerating
						}, (IntPtr)0);
				}
				return _hWndExcel;
			}
		}

		private static object _Application = null;
		public static object Application
		{
			get
			{
				if (_Application == null)
				{
					_Application = GetApplicationFromWindow();
					if (_Application == null)
					{
						// I assume it failed because there was no workbook open
						// Now make workbook with VBA sheet, according to some Google post

						// Create new workbook with the right stuff
						XlCall.Excel(XlCall.xlcEcho, false);
						XlCall.Excel(XlCall.xlcNew, 5);
						XlCall.Excel(XlCall.xlcWorkbookInsert, 6);

						_Application = GetApplicationFromWindow();

						// Clean up
						XlCall.Excel(XlCall.xlcFileClose, false);
						XlCall.Excel(XlCall.xlcEcho, true);
					}
				}
				return _Application;
			}
		}

		private static object GetApplicationFromWindow()
		{
			// This is Andrew Whitechapel's plan for getting the Application object.
			// It does not work when there are no Workbooks open.
			IntPtr hWndMain = WindowHandle;
			IntPtr hWndChild = (IntPtr)0;
			EnumChildWindows(hWndMain, delegate(IntPtr hWndEnum, IntPtr param)
				{
					// Check the window class
					StringBuilder cname = new StringBuilder(256);
					GetClassName(hWndEnum, cname, cname.Capacity);
					if (cname.ToString() == "EXCEL7")
					{
						hWndChild = hWndEnum;
						return false;	// Stop enumerating
					}
					return true;	// Continue enumerating
				} , (IntPtr)0);
			if (hWndChild != (IntPtr)0)
			{
				const uint OBJID_NATIVEOM = 0xFFFFFFF0;
				Guid IID_IDispatch = new Guid(
					 "{00020400-0000-0000-C000-000000000046}");
				IntPtr ptr = (IntPtr)0;
				int hr = AccessibleObjectFromWindow(
						hWndChild, OBJID_NATIVEOM,
						IID_IDispatch.ToByteArray(), ref ptr);
				if (hr >= 0)
				{
					object obj = Marshal.GetObjectForIUnknown(ptr);
					object app = obj.GetType().InvokeMember("Application", System.Reflection.BindingFlags.GetProperty, null, obj, null);
					Marshal.ReleaseComObject(obj);
					//							object ver = app.GetType().InvokeMember("Version", System.Reflection.BindingFlags.GetProperty, null, app, null);
					return app;
				}
			}
			return null;
		}

		public static bool IsInFunctionWizard()
		{
            // TODO: Handle the Find / Find and Replace dialogs.
            //       These should not return true here.
			IntPtr hWndMain = WindowHandle;
			bool inFunctionWizard = false;
			EnumWindows(delegate(IntPtr hWndEnum, IntPtr param)
				{
					// Check the window class
					StringBuilder cname = new StringBuilder(256);
					GetClassName(hWndEnum, cname, cname.Capacity);
					if (cname.ToString().StartsWith("bosa_sdm_XL"))
					{
						if (GetParent(hWndEnum) == hWndMain)
						{
							inFunctionWizard = true;
							return false;	// Stop enumerating
						}
					}
					return true;	// Continue enumerating
				} , (IntPtr)0);

			return inFunctionWizard;
		}
	}
}
