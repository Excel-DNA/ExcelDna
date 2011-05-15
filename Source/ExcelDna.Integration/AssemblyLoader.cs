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
        // Consolidated processing so we only have a single pass through the types.
        // CONSIDER: This is pretty ugly right now (both the flow and the names.)
        //           Try some fancy visitor pattern?
        public static void ProcessAssemblies(
                    List<ExportedAssembly> assemblies,
                    List<MethodInfo> methods,
                    List<ExcelAddInInfo> addIns,
                    List<Type> rtdServerTypes,
                    List<ExcelComClassType> comClassTypes)
        {
            List<AssemblyLoaderExcelServer.ExcelServerInfo> excelServerInfos = new List<AssemblyLoaderExcelServer.ExcelServerInfo>();
            bool loadRibbons = (ExcelDnaUtil.ExcelVersion >= 12.0);

            foreach (ExportedAssembly assembly in assemblies)
            {
                Type[] types = assembly.Assembly.GetTypes();
                bool explicitExports = assembly.ExplicitExports;
                foreach (Type type in types)
                {
                    try
                    {
                        object[] attribs = type.GetCustomAttributes(false);
                        bool isRtdServer;
                        GetExcelMethods(type, explicitExports, methods);
                        AssemblyLoaderExcelServer.GetExcelServerInfos(type, attribs, excelServerInfos);
                        GetExcelAddIns(assembly, type, loadRibbons, addIns);
                        GetRtdServerTypes(type, rtdServerTypes, out isRtdServer);
                        GetComClassTypes(assembly, type, attribs, isRtdServer, comClassTypes);
                    }
                    catch (Exception e)
                    {
                        // TODO: This is unexpected - raise to LogDisplay?
                        Debug.Print("Type {0} could not be analysed. Error: {1}", type.FullName, e.ToString()); 
                    }
                }
            }
            // Sigh. Excel server (service?) stuff is till ugly - but no reeal reason to remove it yet.
            AssemblyLoaderExcelServer.GetExcelServerMethods(excelServerInfos, methods);
        }

        static void GetExcelMethods(Type t, bool explicitExports, List<MethodInfo> excelMethods)
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
                return;
            }

            MethodInfo[] mis = t.GetMethods(BindingFlags.Public | BindingFlags.Static);
            if (explicitExports)
            {
                // Filter list first
                foreach (MethodInfo mi in mis)
                {
                    if (IsMethodMarkedForExport(mi))
                    {
                        excelMethods.Add(mi);
                    }
                }
            }
            else
            {
                // Add all the methods found
                excelMethods.AddRange(mis);
            }
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
            public DnaLibrary ParentDnaLibrary;
        }

		static public void GetExcelAddIns(ExportedAssembly assembly, Type t, bool loadRibbons, List<ExcelAddInInfo> addIns)
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
                    info.ParentDnaLibrary = assembly.ParentDnaLibrary;
                    addIns.Add(info);
                }
            }
            catch (Exception e) // I think only CreateInstance can throw an exception here...
            {
                Debug.Print("GetExcelAddIns CreateInstance problem for type: {0} - exception: {1}", t.FullName, e);
            }

		}

        // DOCUMENT: We register types that implement an interface with the IRtdServer Guid. These include
        //           "Microsoft.Office.Interop.Excel.IRtdServer" and
        //           "ExcelDna.Integration.Rtd.IRtdServer".
        // The RTD server can be accessed using the ExcelDnaUtil.RTD function under the 
        // FullName of the type, or under the ProgId defined in an attribute, if there is one.
        static public void GetRtdServerTypes(Type t, List<Type> rtdServerTypes, out bool isRtdServer)
        {
            isRtdServer = false;
            Type[] itfs = t.GetInterfaces();
            foreach (Type itf in itfs)
            {
                if (itf.GUID == ComAPI.guidIRtdServer)
                {
                    //object[] attrs = t.GetCustomAttributes(typeof(ProgIdAttribute), false);
                    //if (attrs.Length >= 1)
                    //{
                    //    ProgIdAttribute progIdAtt = (ProgIdAttribute)attrs[0];
                    //    rtdServerTypes[progIdAtt.Value] = t;
                    //}
                    //rtdServerTypes[t.FullName] = t;
                    rtdServerTypes.Add(t);
                    isRtdServer = true;
                }
            }
        }

        // DOCUMENT: We register ComVisible types that
        //           (implement IRtdServer OR are in ExternalLibraries marked ComServer='true'
        static public void GetComClassTypes(ExportedAssembly assembly, Type type, object[] attributes, bool isRtdServer, List<ExcelComClassType> comClassTypes)
        {
            if (!Marshal.IsTypeVisibleFromCom(type))
            {
                return;
            }

            if (isRtdServer || assembly.ComServer)
            {
                string progId;
                Guid clsId;

                // Check for public default constructor
                if (type.GetConstructor(BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null) == null)
                {
                    // No use to us here - won't be able to construct.
                    return;
                }

                if (assembly.IsDynamic)
                {
                    // Check that we have an explicit Guid attribute
                    object[] attrs = type.GetCustomAttributes(typeof(GuidAttribute), false);
                    if (attrs.Length == 0)
                    {
                        // No Guid attribute - skip this type.
                        return;
                    }
                    else
                    {
                        GuidAttribute guidAtt = (GuidAttribute)attrs[0];
                        clsId = new Guid(guidAtt.Value);
                    }
                }
                else
                {
                    clsId = Marshal.GenerateGuidForType(type);
                }

                progId = Marshal.GenerateProgIdForType(type);

                ExcelComClassType comClassType = new ExcelComClassType
                {
                    Type = type,
                    ClsId = clsId,
                    ProgId = progId,
                    IsRtdServer = isRtdServer,
                    TypeLibPath = assembly.TypeLibPath
                };
                comClassTypes.Add(comClassType);
            }
        }
	}
}
