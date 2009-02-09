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
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Reflection.Emit;

namespace ExcelDna.Integration
{
	// Loads the managed assembly, finds all the methods to be exported to Excel
	// and build the method information.
	internal class AssemblyLoader
	{
        static public List<MethodInfo> GetExcelMethods(Assembly assembly)
        {
            List<MethodInfo> methods = new List<MethodInfo>();
            Type[] types = assembly.GetTypes();
            foreach (Type t in types)
            {
                // CONSIDER: Implement ExportAll="false" ?
                // DOCUMENT: Exclude if not a class, not public, /*abstract,*/ an array,  
                // open generic type or in "My" namespace.
                // Some basic checks -- what else?
                // TODO: Sort out exactly which types to export.
                if (!t.IsClass || !t.IsPublic ||
                    /*t.IsAbstract ||*/ t.IsArray ||
                    (t.IsGenericType && t.ContainsGenericParameters) ||
                    t.Namespace == "My")
                {
                    // Bad cases
                    Debug.Print("ExcelDNA -> Inappropriate Type: " + t.FullName);
                    continue;
                }

                MethodInfo[] mis = t.GetMethods(BindingFlags.Public | BindingFlags.Static);
                methods.AddRange(mis);
            }

            // This is temporary support for Excel Server
            // TODO: How to make sure this adds no overhead? 
            // Maybe add a new attribute to ExternalLibrary?
            methods.AddRange(AssemblyLoaderExcelServer.GetExcelMethods(assembly));
            return methods;
        }

        // Some support for creating add-ins that are notified of open and close
        // this allows the add-in to add menus, toolbar buttons etc.
        public class ExcelAddInInfo
        {
            public MethodInfo AutoOpenMethod;
            public MethodInfo AutoCloseMethod;
            public object     Instance;
        }

		static public List<ExcelAddInInfo> GetExcelAddIns(Assembly assembly)
		{
			List<ExcelAddInInfo> addIns = new List<ExcelAddInInfo>();
            Type[] types = assembly.GetTypes();
			foreach (Type t in types)
			{
                Type addInType = t.GetInterface("ExcelDna.Integration.IExcelAddIn");
				if (addInType != null)
				{
                    ExcelAddInInfo info = new ExcelAddInInfo();
                    info.AutoOpenMethod = addInType.GetMethod("AutoOpen");
                    info.AutoCloseMethod = addInType.GetMethod("AutoClose");

					info.Instance = Activator.CreateInstance(t);
					addIns.Add(info);
				}
			}
			return addIns;
		}

	}
}
