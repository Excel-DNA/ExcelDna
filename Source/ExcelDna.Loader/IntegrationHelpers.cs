/*
  Copyright (C) 2005-2009 Govert van Drimmelen

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
        static Type integrationType;
        static MethodInfo addCommandMenu;
        static MethodInfo removeCommandMenus;
        
        public static void Bind(Assembly integrationAssembly)
        {
            integrationType = integrationAssembly.GetType("ExcelDna.Integration.Integration");

            Type menuManagerType = integrationAssembly.GetType("ExcelDna.Integration.MenuManager");
            addCommandMenu = menuManagerType.GetMethod("AddCommandMenu", BindingFlags.Static | BindingFlags.NonPublic);
            removeCommandMenus = menuManagerType.GetMethod("RemoveCommandMenus", BindingFlags.Static | BindingFlags.NonPublic);
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
            integrationType.InvokeMember("DnaLibraryAutoOpen", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, null);
        }

        internal static void DnaLibraryAutoClose()
        {
            integrationType.InvokeMember("DnaLibraryAutoClose", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, null);
        }

        internal static string DnaLibraryGetName()
        {
            return (string)integrationType.InvokeMember("DnaLibraryGetName", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, null);
        }

        internal static void InitializeIntegration()
        {
            integrationType.InvokeMember("Initialize", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, null);
        }

        internal static void DeInitializeIntegration()
        {
            integrationType.InvokeMember("DeInitialize", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, null);
        }
    }
}
