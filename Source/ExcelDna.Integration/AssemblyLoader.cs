/*
  Copyright (C) 2005-2010 Govert van Drimmelen

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
using System.Runtime.InteropServices;
using ExcelDna.ComInterop;
using ExcelDna.Integration.CustomUI;

namespace ExcelDna.Integration
{
	// Loads the managed assembly, finds all the methods to be exported to Excel
	// and build the method information.
	internal class AssemblyLoader
	{
        static public List<MethodInfo> GetExcelMethods(ExportedAssembly assembly)
        {
            List<MethodInfo> methods = new List<MethodInfo>();
            Type[] types = assembly.Assembly.GetTypes();
            foreach (Type t in types)
            {
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
				if (assembly.ExplicitExports)
				{
					// Filter list first
					foreach (MethodInfo mi in mis)
					{
						if (IsMethodMarkedForExport(mi))
						{
							methods.Add(mi);
						}
					}
				}
				else
				{
					// Add all the methods found
					methods.AddRange(mis);
				}
            }

            // This is temporary support for Excel Server
            // TODO: How to make sure this adds no overhead? 
            // Maybe add a new attribute to ExternalLibrary?
            methods.AddRange(AssemblyLoaderExcelServer.GetExcelMethods(assembly.Assembly));
            return methods;
        }

		// CAUTION: This check needs to match the usage in ExcelDna.Loader.XlMethodInfo.SetAttributeInfo()
		private static bool IsMethodMarkedForExport(MethodInfo mi)
		{
			object[] atts = mi.GetCustomAttributes(false);
			foreach (object att in atts)
			{
				Type attType = att.GetType();
				if (attType.FullName == "ExcelDna.Integration.ExcelFunctionAttribute" ||
					attType.FullName == "ExcelDna.Integration.ExcelCommandAttribute")
				{
					return true;
				}
			}
			return false;
		}

        // Some support for creating add-ins that are notified of open and close
        // this allows the add-in to add menus, toolbar buttons etc.
        // Also records whether this class should be loaded as a ComAddIn (for the Ribbon).
        public class ExcelAddInInfo
        {
            public MethodInfo AutoOpenMethod;
            public MethodInfo AutoCloseMethod;
            public bool IsCustomUI;
            public object Instance;
        }

		static public List<ExcelAddInInfo> GetExcelAddIns(ExportedAssembly assembly)
		{
			List<ExcelAddInInfo> addIns = new List<ExcelAddInInfo>();
            Type[] types = assembly.Assembly.GetTypes();
            bool loadRibbons = (ExcelDnaUtil.ExcelVersion >= 12.0);

			foreach (Type t in types)
			{
                try
                {
                    
                    Type addInType = t.GetInterface("ExcelDna.Integration.IExcelAddIn");
                    bool isRibbon = (t.BaseType == typeof(ExcelRibbon));
                    if (addInType != null || (isRibbon && loadRibbons) )
                    {
                        ExcelAddInInfo info = new ExcelAddInInfo();
                        if (addInType != null)
                        {
                            info.AutoOpenMethod = addInType.GetMethod("AutoOpen");
                            info.AutoCloseMethod = addInType.GetMethod("AutoClose");
                        }
                        info.IsCustomUI = isRibbon;
                        info.Instance = Activator.CreateInstance(t);
                        addIns.Add(info);
                    }
                }
                catch (Exception e) // I think only CreateInstance can throw an exception here...
                {
                    Debug.Print("GetExcelAddIns CreateInstance problem for type: {0} - exception: {1}", t.FullName, e);
                }
			}
			return addIns;
		}

        // DOCUMENT: We register types that implement an interface with the IRtdServer Guid. These include
        //           "Microsoft.Office.Interop.Excel.IRtdServer" and
        //           "ExcelDna.Integration.Rtd.IRtdServer".
        // The RTD server can be accessed using the ExcelDnaUtil.RTD function under the 
        // FullName of the type, or under the ProgId defined in an attribute, if there is one.
        static public Dictionary<string, Type> GetRtdServerTypes(ExportedAssembly assembly)
        {
			Dictionary<string, Type> rtdServerTypes = new Dictionary<string, Type>();
            Type[] types = assembly.Assembly.GetTypes();
            foreach (Type t in types)
            {
                Type[] itfs = t.GetInterfaces();
                foreach (Type itf in itfs)
                {
                    if (itf.GUID == ComAPI.guidIRtdServer)
                    {
                        object[] attrs = t.GetCustomAttributes(typeof(ProgIdAttribute), false);
                        if (attrs.Length >= 1)
                        {
                            ProgIdAttribute progIdAtt = (ProgIdAttribute)attrs[0];
                            rtdServerTypes[progIdAtt.Value] = t;
                        }
                        rtdServerTypes[t.FullName] = t;
                    }
                }
                //if (t.GetInterface("ExcelDna.Integration.Rtd.IRtdServer") != null ||
                //    t.GetInterface("Microsoft.Office.Interop.Excel.IRtdServer") != null)
                //{
                //}
            }
            return rtdServerTypes;
        }

	}
}
