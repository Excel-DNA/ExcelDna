/*
  Copyright (C) 2005-2011 Govert van Drimmelen

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
using System.Text;

namespace ExcelDna.Integration
{
    /// <summary>
    /// The MenuManager is used by the Loader.
    /// TODO: Can we integrate with ExcelCommandBars?
    ///       Hierarchical menus.
    /// </summary>
  
    // CAUTION: This 'internal' class is called via reflection by the ExcelDna Loader.
    internal static class MenuManager
    {
        internal class MenuEntry
        {
            internal string CommandName;
            internal string MenuName;
            internal string MenuText;

            internal MenuEntry(string commandName, string menuName, string menuText)
            {
                CommandName = commandName;
                MenuName = menuName;
                MenuText = menuText;
            }
        }

        static List<string> addedMenus = new List<string>();
        static List<MenuEntry> addedMenuEntries = new List<MenuEntry>();

        internal static void AddCommandMenu(string commandName, string menuName, string menuText, string description, string shortCut, string helpTopic)
        {
            try // Basically suppress problems here
            {
                bool done = false;
                if (!addedMenus.Contains(menuName))
                {
                    // Check if the menu exists
                    object result = XlCall.Excel(XlCall.xlfGetBar, 1.0 /*Worksheet and Macro sheet*/,
                                                menuName, 0);
                    if (result is ExcelError)
                    {
                        // Add the Menu
                        // DOCUMENT: Description, HelpTopic ???
                        // Throws Access violation exception Excel if I add a string to description or helptopic
                        XlCall.Excel(XlCall.xlfAddMenu, 1.0 /*Worksheet and Macro sheet*/,
                            new object[,] { { menuName, null, null, null, null},
				                            { menuText, commandName, 
												null/*shortcut_key (Mac Only)*/, 
												null, // mi.Description, 
												null /*mi.HelpTopic*/} });
                        addedMenus.Add(menuName);
                        done = true;
                    }
                }
                if (!done)
                {
                    // Check if command exists
                    object result = XlCall.Excel(XlCall.xlfGetBar, 1.0 /*Worksheet and Macro sheet*/,
                                                menuName, menuText);
                    if (result is ExcelError)
                    {
                        // Register the new command on the menu
                        XlCall.Excel(XlCall.xlfAddCommand, 
                                     1.0 /*Worksheet and Macro sheet*/, 
                                     menuName,
                                     new object[] { 
                                        menuText, 
                                        commandName, 
										null/*shortcut_key (Mac Only)*/, 
										null, // mi.Description, 
										null /*mi.HelpTopic*/});
                        addedMenuEntries.Add(new MenuEntry(commandName, menuName, menuText));
                    }
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }

        internal static void RemoveCommandMenus()
        {
            // First take out menus and commands
            foreach (MenuEntry me in addedMenuEntries)
            {
                try
                {
                    XlCall.Excel(XlCall.xlfDeleteCommand, 
                                 1.0 /*Worksheet and Macro sheet*/,
                                 me.MenuName, 
                                 me.MenuText);
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
            }
            addedMenuEntries.Clear();

            foreach (string menu in addedMenus)
            {
                try
                {
                    XlCall.Excel(XlCall.xlfDeleteMenu, 
                                 1.0 /*Worksheet and Macro sheet*/,
                                 menu);
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
            }
            addedMenus.Clear();
        }
    }
}
