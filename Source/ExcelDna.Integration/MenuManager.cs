//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Logging;

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
        static IMenuManager _menuManager;
        static MenuManager()
        {
            if (ExcelDnaUtil.SafeIsExcelVersionPre15)
            {
                _menuManager = new ExcelPre15MenuManager();
            }
            else
            {
                _menuManager = new Excel15MenuManager();
            }
        }

        // These methods are called from XlRegistration via reflection 
        // The binding is in IntegrationHelpers
        static void AddCommandMenu(string commandName, string menuName, string menuText, string description, string shortCut, string helpTopic)
        {
            _menuManager.AddCommandMenu(commandName, menuName, menuText, description, shortCut, helpTopic);
        }

        static void RemoveCommandMenus()
        {
            _menuManager.RemoveCommandMenus();
        }
    }

    interface IMenuManager
    {
        void AddCommandMenu(string commandName, string menuName, string menuText, string description, string shortCut, string helpTopic);
        void RemoveCommandMenus();
    }

    class ExcelPre15MenuManager : IMenuManager
    {
        private class MenuEntry
        {
            internal readonly string CommandName;
            internal readonly string MenuName;
            internal readonly string MenuText;

            internal MenuEntry(string commandName, string menuName, string menuText)
            {
                CommandName = commandName;
                MenuName = menuName;
                MenuText = menuText;
            }
        }

        readonly List<string> _addedMenus = new List<string>();
        readonly List<MenuEntry> _addedMenuEntries = new List<MenuEntry>();

        public void AddCommandMenu(string commandName, string menuName, string menuText, string description, string shortCut, string helpTopic)
        {
            try // Basically suppress problems here..?
            {
                bool done = false;
                if (!_addedMenus.Contains(menuName))
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
                        _addedMenus.Add(menuName);
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
                        _addedMenuEntries.Add(new MenuEntry(commandName, menuName, menuText));
                    }
                }
            }
            catch (Exception e)
            {
                Logger.Initialization.Error(e, "MenuManager.AddCommandMenu Error");
            }
        }

        public void RemoveCommandMenus()
        {
            // First take out menus and commands
            foreach (MenuEntry me in _addedMenuEntries)
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
                    Logger.Initialization.Error(e, "MenuManager.RemoveCommandMenus Error");
                }
            }
            _addedMenuEntries.Clear();

            foreach (string menu in _addedMenus)
            {
                try
                {
                    XlCall.Excel(XlCall.xlfDeleteMenu,
                                 1.0 /*Worksheet and Macro sheet*/,
                                 menu);
                }
                catch (Exception e)
                {
                    Logger.Initialization.Error(e, "MenuManager.RemoveCommandMenus Error");
                }
            }
            _addedMenus.Clear();
        }
    }

    class Excel15MenuManager : IMenuManager
    {
        readonly Dictionary<string, CommandBarPopup> _foundMenus = new Dictionary<string, CommandBarPopup>();
        readonly List<CommandBarPopup> _addedMenus = new List<CommandBarPopup>();
        readonly List<CommandBarButton> _addedButtons = new List<CommandBarButton>();

        public void AddCommandMenu(string commandName, string menuName, string menuText, string description, string shortCut, string helpTopic)
        {
            try // Basically suppress problems here..?
            {
                CommandBarPopup menu;
                if (!_foundMenus.TryGetValue(menuName, out menu))
                {
                    // We've not seen this menu before
                    
                    // Check if the menu exists
                    CommandBars commandBars = ExcelCommandBarUtil.GetCommandBars();
                    CommandBar worksheetBar = commandBars[1];
                    CommandBarControls controls = worksheetBar.Controls;
                    int controlCount = controls.Count();

                    for (int i = 1; i <= controlCount; i++)
                    {
                        CommandBarControl control = controls[i];
                        if (control.Caption == menuName && control is CommandBarPopup)
                        {
                            menu = (CommandBarPopup)control;
                            _foundMenus[menuName] = menu;
                            break;
                        }
                    }

                    if (menu == null)
                    {
                        // Make a new menu
                        menu = controls.AddPopup(menuName);
                        menu.Caption = menuName;
                        _addedMenus.Add(menu);
                        _foundMenus[menuName] = menu;
                    }
                }

                CommandBarControls menuButtons = menu.Controls;
                int buttonCount = menu.Controls.Count();
                for (int i = 1; i <= buttonCount; i++)
                {
                    CommandBarControl button = menuButtons[i];
                    if (button.Caption == menuText && button is CommandBarButton)
                    {
                        button.OnAction = commandName;
                        return;
                    }
                }

                // If we're here, need to add a button.
                CommandBarButton newButton = menuButtons.AddButton();
                newButton.Caption = menuText;
                newButton.OnAction = commandName;
                _addedButtons.Add(newButton);
            }
            catch (Exception e)
            {
                Logger.Initialization.Error(e, "MenuManager.AddCommandMenu Error");
            }
        }

        public void RemoveCommandMenus()
        {
            foreach (CommandBarButton button in _addedButtons)
            {
                button.Delete(true);
            }

            foreach (CommandBarPopup popup in _addedMenus)
            {
                popup.Delete(true);
            }
        }
    }
}
