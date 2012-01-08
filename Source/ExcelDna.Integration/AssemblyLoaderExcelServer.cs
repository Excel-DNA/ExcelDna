/*
  Copyright (C) 2005-2012 Govert van Drimmelen

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
using System.Reflection;
using System.Reflection.Emit;

// First attempt at support for Excel Services.

// The information I use is from http://blogs.msdn.com/cumgranosalis/default.aspx
// The Excel Services managed Udf model allows classes to be marked with
// Microsoft.Office.Excel.Server.Udf.UdfClass
// and some of the class's methods to be marked as Udfs using 
// Microsoft.Office.Excel.Server.Udf.UdfMethod.
// Previously ExcelDna supported only static methods to be exported, 
// and the main issue for compatibility is to decide how to handle the object 
// instantiation.
// (This should not be confused with proper support for object references as 
// described by the Excel gurus.
// Built-in support for this might come later to ExcelDna 
// but it could be currently implemented in a user library without 
// further built-in support, just as it would be done inside an .xll.)
// In Excel Services the objects are instantiated per session,
// in the client model it's not so clear what to do.
// I think the ideal plan is to lazy-instantiate the objects, but this would
// require the generation of stubs that handle the checking and instantiation.
// So my first-pass solution is to create all the objects when the library is 
// first loaded, and keep them around forever.

// CONSIDER: The support for instance methods actually impacts quite a lot of the XlMethodInfo story.
// So the right way to add this support would be to refactor XlMethodInfo into a hierarchy
// with children for functions and macros, and for instance methods vs. static methods.

namespace ExcelDna.Integration
{
	// Loads the managed assembly, finds all the methods to be exported to Excel
	// and build the method information.
    internal class AssemblyLoaderExcelServer
    {

        internal class ExcelServerInfo
        {
            public object Instance;
            public List<MethodInfo> Methods;
        }

        internal static void GetExcelServerInfos(Type t, object[] attribs, List<ExcelServerInfo> excelServerInfos)
        {
            bool isUdfClass = false;
            foreach (object attrib in attribs)
            {
                Type attribType = attrib.GetType();
                if (attribType.FullName == "Microsoft.Office.Excel.Server.Udf.UdfClassAttribute")
                {
                    // Candidate for export
                    isUdfClass = true;
                    break;
                }
            }
            if (!isUdfClass) return;

            // Instantiate -- the class must have a parameterless constructor
            ConstructorInfo ci = t.GetConstructor(new Type[0]);
            if (ci == null)
            {
                // Bad case
                Debug.Print("ExcelDNA -> UdfClass: " + t.FullName + " has no parameterless constructor.");
                return;
            }

            try
            {
                object instance = ci.Invoke(null);
                ExcelServerInfo serverInfo = new ExcelServerInfo();
                excelServerInfos.Add(serverInfo);
                serverInfo.Instance = instance;
                serverInfo.Methods = new List<MethodInfo>();

                // Now go through all the methods, finding those with the UdfMethod attribute
                foreach (MethodInfo method in t.GetMethods(BindingFlags.Public | BindingFlags.Instance))
                {
                    // Simple check that this is a function
                    if (method.ReturnType == typeof(void))
                    {
                        // Bad case
                        Debug.Print("ExcelDNA -> UdfMethod: " + method.Name + " returns void.");
                        continue;
                    }

                    foreach (object attrib in method.GetCustomAttributes(false))
                    {
                        // CONSIDER: Does this GetType() require that the assembly which defines
                        // the attribute be available?
                        Type attribType = attrib.GetType();

                        if (attribType.FullName == "Microsoft.Office.Excel.Server.Udf.UdfMethodAttribute")
                        {
                            // Candidate for export
                            serverInfo.Methods.Add(method);
                        }
                    }
                }
            }
            catch { }
        }

        static internal void GetExcelServerMethods(List<ExcelServerInfo> serverInfos, List<MethodInfo> methods)
        {

            // Prevent dynamic assembly if no Udf* methods
            if (serverInfos.Count == 0) return;

            // Now we build a module with 
            AssemblyBuilder assemblyBuilder;
            ModuleBuilder moduleBuilder;
            assemblyBuilder = AppDomain.CurrentDomain.DefineDynamicAssembly(
                new AssemblyName("ExcelDna.ExcelServer.DynamicDelegateAssembly"/*TODO:Add assembly name to disambiguate different add-ins?*/),
                AssemblyBuilderAccess.Run/*AndSave*/);
            moduleBuilder = assemblyBuilder.DefineDynamicModule("DynamicDelegates");
            TypeBuilder typeBuilder = moduleBuilder.DefineType("DynamicType");

            // First pass - create all the static fields and redirection code
            foreach (ExcelServerInfo serverInfo in serverInfos)
            {
                string fieldName = serverInfo.Instance.GetType().Name;
                FieldBuilder fieldBuilder = typeBuilder.DefineField(fieldName, serverInfo.Instance.GetType(), FieldAttributes.Public | FieldAttributes.Static);
                foreach (MethodInfo method in serverInfo.Methods)
                {
                    // Add a static method to the type that contains an instance
                    // and a call to the instance method.
                    ParameterInfo[] parameterInfos = method.GetParameters();
                    Type[] paramTypes = Array.ConvertAll<ParameterInfo, Type>(parameterInfos, delegate(ParameterInfo parameterInfo) { return parameterInfo.ParameterType; });
                    MethodBuilder methodBuilder = typeBuilder.DefineMethod(method.Name,
                                    MethodAttributes.Public | MethodAttributes.Static,
                                    method.ReturnType,
                                    paramTypes);

                    // TODO: Set attributes on parameters so that custom function categories etc. will work
                    ILGenerator wrapIL = methodBuilder.GetILGenerator();
                    LocalBuilder retobj = null;
                    // Make a local to contain the return value
                    retobj = wrapIL.DeclareLocal(method.ReturnType);

                    wrapIL.Emit(OpCodes.Ldsfld, fieldBuilder);
                    // push all the arguments
                    for (byte i = 0; i < paramTypes.Length; i++)
                    {
                        wrapIL.Emit(OpCodes.Ldarg_S, i);
                    }
                    // Call the real method
                    wrapIL.EmitCall(OpCodes.Call, method, null);
                    //// Store the return value into the local variable
                    wrapIL.Emit(OpCodes.Stloc_S, retobj);
                    //// Push the return value
                    wrapIL.Emit(OpCodes.Ldloc_S, retobj);
                    wrapIL.Emit(OpCodes.Ret);

                }
            }

            // Bake the type
            Type wrapperType = typeBuilder.CreateType();

            // Second pass - set field values and create delegates
            foreach (ExcelServerInfo serverInfo in serverInfos)
            {
                string typeName = serverInfo.Instance.GetType().Name;
                FieldInfo field = wrapperType.GetField(typeName);
                field.SetValue(null, serverInfo.Instance);
                foreach (MethodInfo method in serverInfo.Methods)
                {
                    MethodInfo wrapperMethod = wrapperType.GetMethod(method.Name);
                    methods.Add(wrapperMethod);
                }
            }
        }
    }
}
