/*
  Copyright (C) 2005, 2006, 2007 Govert van Drimmelen

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
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelDna.Integration
{
	internal delegate void SetJumpDelegate(int fi, IntPtr pfn);

	// Implements an XLL in managed code
	public class XlLibrary 
	{
		static SetJumpDelegate setJump;
		static List<XlMethodInfo> registeredMethods = new List<XlMethodInfo>();
		static List<string> addedMenus = new List<string>();
		static List<XlMethodInfo> addedCommands = new List<XlMethodInfo>();
		static string dllName;

		internal static SetJumpDelegate SetJump
		{
			set { setJump = value; }
		}

		// TODO: Improve the separation between the Xll registration and the 
		// assemblies being registered.
		// Functions that an XlLibrary must implement
		internal static short AutoOpen()
		{
			short result = 0;
			try
			{
				// Clear any references, if we are already loaded
				UnregisterMethods();

				// Initialize loads the .dna file
				DnaLibrary.Initialize();
				dllName = Assembly.GetExecutingAssembly().Location;
				XlCall.Excel(XlCall.xlcMessage, true, "Registering library " + dllName);

				try
				{
					// Ensure there is an Application object created
					object temp = Excel.Application;
				}
				catch (Exception e)
				{
					// TODO: What to do here?
					Debug.WriteLine(e.Message);
				}

                DnaLibrary.CurrentLibrary.AutoOpen();

				result = 1; // All is OK
			}
			catch (Exception e)
			{
				// TODO: What to do here?
				Debug.WriteLine(e.Message);
				result = 0;
			}
			finally
			{
				// Clear the status bar message
				XlCall.Excel(XlCall.xlcMessage, false);
			}

			return result;
		}

		internal static void AutoFree(IntPtr pXloper)
		{
			// CONSIDER: This might be improved....
			// Another option would be to have the Com memory allocator run in unmanaged code.
			// Right now I think this is OK, and easiest from where I'm coming.
			// This function can only be called after a return from a user function.
			// I just free all the possibly big memory allocations.

			XlObjectArrayMarshaler.FreeMemory();
		}
        
        public static void RegisterMethods(List<MethodInfo> methods)
		{
            List<XlMethodInfo> xlMethods = XlMethodInfo.ConvertToXlMethodInfos(methods);

            xlMethods.ForEach(RegisterXlMethod);
		}

        private static void RegisterXlMethod(XlMethodInfo mi)
        {
            // TODO: Store the handle (but no unregistration for now)
            int index = registeredMethods.Count;
            setJump(index, mi.FunctionPointer);
            String procName = String.Format("f{0}", index);

            string functionType = mi.ReturnType == null ? "" : mi.ReturnType.XlType.ToString();
            string argumentNames = "";
            bool showDescriptions = false;
            string[] argumentDescriptions = new string[mi.Parameters.Length];

            for (int j = 0; j < mi.Parameters.Length; j++)
            {
                XlParameterInfo pi = mi.Parameters[j];

                functionType += pi.XlType;
                if (j > 0)
                    argumentNames += ", ";
                argumentNames += pi.Name;
                argumentDescriptions[j] = pi.Description;

                if (pi.Description != "")
                    showDescriptions = true;

                // DOCUMENT: Here is the patch for the Excel Function Description bug.
                // DOCUMENT: I add ". " to the last parameters.
                if (j == mi.Parameters.Length - 1)
                    argumentDescriptions[j] += ". ";

            } // for each parameter

            if (mi.IsVolatile)
                functionType += "!";
            // TODO: How do these interact ?
            // DOCUMENT: If # is set and there is an R argument, 
            // Excel considers the function volatile
            // You can call xlfVolatile, false in beginning of function to clear.
            if (mi.IsMacroType)
                functionType += "#";

            // DOCUMENT: Here is the patch for the Excel Function Description bug.
            // DOCUMENT: I add ". " if the function takes no parameters.
            string functionDescription = mi.Description;
            if (mi.Parameters.Length == 0)
                functionDescription += ". ";

            if (mi.Description != "")
                showDescriptions = true;

            // DOCUMENT: When there is no description, we don't add any.
            // This allows the user to work around the Excel bug where an extra parameter is displayed if
            // the function has no parameter but displays a description
            int numArguments;
            // DOCUMENT: Maximum 20 Argument Descriptions when registering using Excel4 function.
            int numArgumentDescriptions;
            if (showDescriptions)
            {
                numArgumentDescriptions = Math.Min(argumentDescriptions.Length, 20);
                numArguments = 10 + numArgumentDescriptions;
            }
            else
            {
                numArgumentDescriptions = 0;
                numArguments = 9;
            }

            object[] registerParameters = new object[numArguments];
            registerParameters[0] = dllName;
            registerParameters[1] = procName;
            registerParameters[2] = functionType;
            registerParameters[3] = mi.Name;
            registerParameters[4] = argumentNames;
            registerParameters[5] = mi.IsCommand ? 2 /*macro*/
                                                          : (mi.IsHidden ? 0 : 1); /*function*/
            registerParameters[6] = mi.Category;
            registerParameters[7] = mi.ShortCut; /*shortcut_text*/
            registerParameters[8] = mi.HelpTopic; /*help_topic*/ ;

            if (showDescriptions)
            {
                registerParameters[9] = functionDescription;

                for (int k = 0; k < numArgumentDescriptions; k++)
                {
                    registerParameters[10 + k] = argumentDescriptions[k];
                }
            }

            // Basically suppress problems here !?
            try
            {
                mi.RegisterId = (double)XlCall.Excel(XlCall.xlfRegister, registerParameters);
                registeredMethods.Add(mi);
            }
            catch (Exception e)
            {
                // TODO: What to do here?
                Debug.WriteLine(e.Message);
            }

            // CONSIDER: The menu stuff might fit better elsewhere?
            if (mi.IsCommand
                && mi.MenuName != null && mi.MenuName != ""
                && mi.MenuText != null && mi.MenuText != "")
            {
                RegisterMenu(mi);
            }
        }

		private static void RegisterMenu(XlMethodInfo mi)
		{
			try // Basically suppress problems here
			{
				bool done = false;
				if (!addedMenus.Contains(mi.MenuName))
				{
					// Check if the menu exists
					object result = XlCall.Excel(XlCall.xlfGetBar, 1.0 /*Worksheet and Macro sheet*/,
												mi.MenuName, 0);
					if (result is ExcelError)
					{
						// Add the Menu
						// DOCUMENT: Description, HelpTopic ???
						// Throws Access violation exception Excel if I add a string to description or helptopic
						XlCall.Excel(XlCall.xlfAddMenu, 1.0 /*Worksheet and Macro sheet*/,
							new object[,] { { mi.MenuName, null, null, null, null},
				                            { mi.MenuText, mi.Name, 
												null/*shortcut_key (Mac Only)*/, 
												null, // mi.Description, 
												null /*mi.HelpTopic*/} });
						addedMenus.Add(mi.MenuName);
						done = true;
					}
				}
				if (!done)
				{
					// Check if command exists
					object result = XlCall.Excel(XlCall.xlfGetBar, 1.0 /*Worksheet and Macro sheet*/,
												mi.MenuName, mi.MenuText);
					if (result is ExcelError)
					{
						// Register the new command on the menu
						XlCall.Excel(XlCall.xlfAddCommand, 1.0 /*Worksheet and Macro sheet*/, mi.MenuName,
							new object[] { mi.MenuText, mi.Name, 
												null/*shortcut_key (Mac Only)*/, 
												null, // mi.Description, 
												null /*mi.HelpTopic*/});
						addedCommands.Add(mi);
					}
				}
			}
			catch (Exception e)
			{
				Debug.WriteLine(e.Message);
			}
		}

		private static void UnregisterMethods()
		{
			// First take out menus and commands
			foreach (XlMethodInfo mi in addedCommands)
			{
				try
				{
					XlCall.Excel(XlCall.xlfDeleteCommand, 1.0 /*Worksheet and Macro sheet*/,
						mi.MenuName, mi.MenuText);
				}
				catch (Exception e)
				{
					Debug.WriteLine(e.Message);
				}
			}
			addedCommands.Clear();
			foreach (string menu in addedMenus)
			{
				try
				{
					XlCall.Excel(XlCall.xlfDeleteMenu, 1.0 /*Worksheet and Macro sheet*/,
						menu);
				}
				catch (Exception e)
				{
					Debug.WriteLine(e.Message);
				}
			}
			addedMenus.Clear();

			// Now take out the methods
			foreach (XlMethodInfo mi in registeredMethods)
			{
				try
				{
					if (mi.IsCommand)
					{
						XlCall.Excel(XlCall.xlfSetName, mi.Name, "");
					}
					else
					{
						// I follow the advice from X-Cell website
						// to get function out of Wizard
						XlCall.Excel(XlCall.xlfRegister, dllName, "xlAutoRemove", "J", mi.Name, Missing.Value, 0);
					}
					XlCall.Excel(XlCall.xlfUnregister, mi.RegisterId);
				}
				catch (Exception e)
				{
					// TODO: What to do here?
					Debug.WriteLine(e.Message);
				}
			}
			registeredMethods.Clear();
		}

		internal static short AutoClose()
		{
            DnaLibrary.CurrentLibrary.AutoClose();
			// UnregisterMethods();
			return 1; // 0 if problems ?
		}

		internal static short AutoAdd()
		{
			return 1; // 0 if problems ?
		}

		internal static short AutoRemove()
		{
			// Apparently better if called here, 
			// so I try to, but make it safe to call again.
			UnregisterMethods();
			return 1; // 0 if problems ?
		}

		internal static IntPtr AddInManagerInfo(IntPtr pXloperAction)
		{
			ICustomMarshaler m = XlObjectMarshaler.GetInstance("");
			object action = m.MarshalNativeToManaged(pXloperAction);
			object result;
            if ((action is short && (short)action == 1) ||
                (action is double && (double)action == 1))
            {
               result = DnaLibrary.CurrentLibrary.Name;
            }
            else
                result = ExcelError.ExcelErrorValue;
			return m.MarshalManagedToNative(result);
		}
	}
}
