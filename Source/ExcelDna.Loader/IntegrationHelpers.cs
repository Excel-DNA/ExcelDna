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
using System.Text;
using System.Reflection;

namespace ExcelDna.Loader
{
    // This class has a friend in XlCustomMarshal.
    internal static class IntegrationHelpers
    {
        static MethodInfo addCommandMenu;
        static MethodInfo removeCommandMenus;
        static object dnaLibraryCurrentLibrary;
        static MethodInfo dnaLibraryAutoOpen;
        static MethodInfo dnaLibraryAutoClose;
        static PropertyInfo dnaLibraryGetName;

        public static void Initialize(Assembly integrationAssembly)
        {
            IntegrationMarshalHelpers.Initialize(integrationAssembly);
            Type menuManagerType = integrationAssembly.GetType("ExcelDna.Integration.MenuManager");
            addCommandMenu = menuManagerType.GetMethod("AddCommandMenu", BindingFlags.Static | BindingFlags.NonPublic);
            removeCommandMenus = menuManagerType.GetMethod("RemoveCommandMenus", BindingFlags.Static | BindingFlags.NonPublic);

            Type dnaLibraryType = integrationAssembly.GetType("ExcelDna.Integration.DnaLibrary");
            dnaLibraryCurrentLibrary = dnaLibraryType.GetProperty("CurrentLibrary",  BindingFlags.Static | BindingFlags.Public).GetValue(null, null);

            dnaLibraryAutoOpen = dnaLibraryType.GetMethod("AutoOpen", BindingFlags.Instance | BindingFlags.NonPublic);
            dnaLibraryAutoClose = dnaLibraryType.GetMethod("AutoClose", BindingFlags.Instance | BindingFlags.NonPublic);

            dnaLibraryGetName = dnaLibraryType.GetProperty("Name");
        }

        public static void AddCommandMenu(string commandName, string menuName, string menuText, string description, string shortCut, string helpTopic)
        {
            addCommandMenu.Invoke(null, new object[] { commandName, menuName, menuText, description, shortCut, helpTopic});
        }

        public static void RemoveCommandMenus()
        {
            removeCommandMenus.Invoke(null, null);
        }

        internal static void DnaLibraryAutoOpen()
        {
            dnaLibraryAutoOpen.Invoke(dnaLibraryCurrentLibrary, null);
        }

        internal static void DnaLibraryAutoClose()
        {
            dnaLibraryAutoClose.Invoke(dnaLibraryCurrentLibrary, null);
        }

        internal static string DnaLibraryGetName()
        {
            return (string)dnaLibraryGetName.GetValue(dnaLibraryCurrentLibrary, null);
        }
    }
}
