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
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;

namespace ExcelDna.Loader
{
	// TODO: Refactor into XlFunctionInfo and XlCommandInfo ?

	internal class XlMethodInfo
	{
        public static int Index = 0;

		public GCHandle DelegateHandle; // TODO: What with this - when to clean up? 
										// For cleanup should call DelegateHandle.Free()
		public IntPtr FunctionPointer;

		// Info for Excel Registration
		public bool   IsCommand;
		public string Name;			// Name of UDF/Macro in Excel
		public string Description;
		public bool   IsHidden;		// For Functions only
		public string ShortCut;		// For macros only
		public string MenuName;		// For macros only
		public string MenuText;     // For macros only
		public string Category;
		public bool   IsVolatile;
		public bool   IsExceptionSafe;
		public bool   IsMacroType;
        public bool   IsThreadSafe; // For Functions only
        public bool   IsClusterSafe;// For Functions only
		public string HelpTopic;
		public double RegisterId;

		public XlParameterInfo[] Parameters;
		public XlParameterInfo ReturnType; // Macro will have ReturnType null

		// THROWS: Throws a DnaMarshalException if the method cannot be turned into an XlMethodInfo
		// TODO: Manage errors if things go wrong
		private XlMethodInfo(MethodInfo targetMethod, ModuleBuilder modBuilder)
		{
			// Default Name, Description and Category
			Name = targetMethod.Name;
			Description = "";
            Category = IntegrationHelpers.DnaLibraryGetName();
			HelpTopic = "";
			IsVolatile = false;
			IsExceptionSafe = false;
			IsHidden = false;
			IsMacroType = false;
            IsThreadSafe = false;
            IsClusterSafe = false;

			ShortCut = "";
			// DOCUMENT: Default MenuName is the library name
			// but menu is only added if at least the MenuText is set.
            MenuName = IntegrationHelpers.DnaLibraryGetName();
			MenuText = null;	// Menu is only 

            SetAttributeInfo(targetMethod.GetCustomAttributes(false));
            
            // Return type conversion
			if (targetMethod.ReturnType == typeof(void))
			{
				IsCommand = true;
				ReturnType = null;
			}
			else
			{
				IsCommand = false;
				ReturnType = new XlParameterInfo(targetMethod.ReturnType, true, IsExceptionSafe);
			}

			// Parameters - meta-data and type conversion
			Parameters = 
					Array.ConvertAll<ParameterInfo, XlParameterInfo>( 
					targetMethod.GetParameters(),
					delegate(ParameterInfo pi) { return new XlParameterInfo(pi); });

			// Create the delegate type, wrap the targetMethod and create the delegate
			Type delegateType = CreateDelegateType(modBuilder);
			Delegate xlDelegate = CreateMethodDelegate(targetMethod, delegateType);

			// Need to add a reference to prevent garbage collection of our delegate
			// Don't need to pin, according to 
			// "How to: Marshal Callbacks and Delegates Using C++ Interop"
			// Currently this delegate is never released
			// TODO: Clean up properly
			DelegateHandle = GCHandle.Alloc(xlDelegate);
			FunctionPointer = Marshal.GetFunctionPointerForDelegate(xlDelegate);
		}

		// Basic setup - get description, category etc.
		private void SetAttributeInfo(object[] attributes)
		{
			// DOCUMENT: Description in ExcelFunctionAtribute overrides DescriptionAttribute
			// DOCUMENT: Default Category is Current Library Name.
			// Get System.ComponentModel.DescriptionAttribute
			// Search through attribs for Description
			foreach (object attrib in attributes)
			{
				System.ComponentModel.DescriptionAttribute desc =
					attrib as System.ComponentModel.DescriptionAttribute;
				if (desc != null)
				{
					Description = desc.Description;
				}

                // There was a problem with the type identification when checking the 
                // attribute types, for the second instance of the .xll 
                // that is loaded.
                // So I check on the names and access through reflection.
                // CONSIDER: Fix again? It should rather be 
                //ExcelFunctionAttribute xlfunc = attrib as ExcelFunctionAttribute;
                //if (xlfunc != null)
                //{
                //    if (xlfunc.Name != null)
                //    {
                //        Name = xlfunc.Name;
                //    }
                //    if (xlfunc.Description != null)
                //    {
                //        Description = xlfunc.Description;
                //    }
                //    if (xlfunc.Category != null)
                //    {
                //        Category = xlfunc.Category;
                //    }
                //    if (xlfunc.HelpTopic != null)
                //    {
                //        HelpTopic = xlfunc.HelpTopic;
                //    }
                //    IsVolatile = xlfunc.IsVolatile;
                //    IsExceptionSafe = xlfunc.IsExceptionSafe;
                //    IsMacroType = xlfunc.IsMacroType;
                //}
                //ExcelCommandAttribute xlcmd = attrib as ExcelCommandAttribute;
                //if (xlcmd != null)
                //{
                //    if (xlcmd.Name != null)
                //    {
                //        Name = xlcmd.Name;
                //    }
                //    if (xlcmd.Description != null)
                //    {
                //        Description = xlcmd.Description;
                //    }
                //    if (xlcmd.HelpTopic != null)
                //    {
                //        HelpTopic = xlcmd.HelpTopic;
                //    }
                //    if (xlcmd.ShortCut != null)
                //    {
                //        ShortCut = xlcmd.ShortCut;
                //    }
                //    if (xlcmd.MenuName != null)
                //    {
                //        MenuName = xlcmd.MenuName;
                //    }
                //    if (xlcmd.MenuText != null)
                //    {
                //        MenuText = xlcmd.MenuText;
                //    }
                //    IsHidden = xlcmd.IsHidden;
                //    IsExceptionSafe = xlcmd.IsExceptionSafe;
                //}
                
                Type attribType = attrib.GetType();
                if (attribType.FullName == "ExcelDna.Integration.ExcelFunctionAttribute")
                {
                    string name = (string)attribType.GetField("Name").GetValue(attrib);
                    string description = (string)attribType.GetField("Description").GetValue(attrib);
                    string category = (string)attribType.GetField("Category").GetValue(attrib);
                    string helpTopic = (string)attribType.GetField("HelpTopic").GetValue(attrib);
                    bool isVolatile = (bool)attribType.GetField("IsVolatile").GetValue(attrib);
                    bool isExceptionSafe = (bool)attribType.GetField("IsExceptionSafe").GetValue(attrib);
                    bool isMacroType = (bool)attribType.GetField("IsMacroType").GetValue(attrib);
                    bool isHidden = (bool)attribType.GetField("IsHidden").GetValue(attrib);
                    bool isThreadSafe = (bool)attribType.GetField("IsThreadSafe").GetValue(attrib);
                    bool isClusterSafe = (bool)attribType.GetField("IsClusterSafe").GetValue(attrib);
                    if (name != null)
                    {
                        Name = name;
                    }
                    if (description != null)
                    {
                        Description = description;
                    }
                    if (category != null)
                    {
                        Category = category;
                    }
                    if (helpTopic != null)
                    {
                        HelpTopic = helpTopic;
                    }
                    IsVolatile = isVolatile;
                    IsExceptionSafe = isExceptionSafe;
                    IsMacroType = isMacroType;
                    IsHidden = isHidden;
                    IsThreadSafe = (!isMacroType && isThreadSafe);
                    // DOCUMENT: IsClusterSafe function MUST NOT be marked as IsMacroType=true and MAY be marked as IsThreadSafe = true.
                    //           [xlfRegister (Form 1) page in the Microsoft Excel 2010 XLL SDK Documentation]
                    IsClusterSafe = (!isMacroType && isClusterSafe);
                }

                if (attribType.FullName == "ExcelDna.Integration.ExcelCommandAttribute")
                {
                    string name = (string)attribType.GetField("Name").GetValue(attrib);
                    string description = (string)attribType.GetField("Description").GetValue(attrib);
                    string helpTopic = (string)attribType.GetField("HelpTopic").GetValue(attrib);
                    string shortCut = (string)attribType.GetField("ShortCut").GetValue(attrib);
                    string menuName = (string)attribType.GetField("MenuName").GetValue(attrib);
                    string menuText = (string)attribType.GetField("MenuText").GetValue(attrib);
//                    bool isHidden = (bool)attribType.GetField("IsHidden").GetValue(attrib);
                    bool isExceptionSafe = (bool)attribType.GetField("IsExceptionSafe").GetValue(attrib);

                    if (name != null)
                    {
                        Name = name;
                    }
                    if (description != null)
                    {
                        Description = description;
                    }
                    if (helpTopic != null)
                    {
                        HelpTopic = helpTopic;
                    }
                    if (shortCut != null)
                    {
                        ShortCut = shortCut;
                    }
                    if (menuName != null)
                    {
                        MenuName = menuName;
                    }
                    if (menuText != null)
                    {
                        MenuText = menuText;
                    }
//                    IsHidden = isHidden;  // Only for functions.
                    IsExceptionSafe = isExceptionSafe;
                }
			}
		}

		private Type CreateDelegateType(ModuleBuilder modBuilder)
		{
			TypeBuilder typeBuilder;
			MethodBuilder methodBuilder;
			XlParameterInfo[] paramInfos = Parameters;
			Type[] paramTypes = Array.ConvertAll<XlParameterInfo, Type>( paramInfos,
				delegate(XlParameterInfo pi) { return pi.DelegateParamType; });

			// Create a delegate that has the same signature as the method we would like to hook up to
			typeBuilder = modBuilder.DefineType("f" + Index++ + "Delegate",
							TypeAttributes.Class  | TypeAttributes.Public | 
							TypeAttributes.Sealed,
							typeof(System.MulticastDelegate));
			ConstructorBuilder constructorBuilder = typeBuilder.DefineConstructor(
							MethodAttributes.RTSpecialName | MethodAttributes.HideBySig | 
							MethodAttributes.Public, CallingConventions.Standard,
							new Type[] { typeof(object), typeof(int) });
			constructorBuilder.SetImplementationFlags(MethodImplAttributes.Runtime |
													  MethodImplAttributes.Managed );

			// Build up the delegate
			// Define the Invoke method for the delegate
			methodBuilder = typeBuilder.DefineMethod("Invoke",
							MethodAttributes.Public | MethodAttributes.HideBySig |
							MethodAttributes.NewSlot | MethodAttributes.Virtual,
							IsCommand ? typeof(void) : ReturnType.DelegateParamType, // What here for macro? null or Void ?
							paramTypes);
			methodBuilder.SetImplementationFlags(MethodImplAttributes.Runtime |
							MethodImplAttributes.Managed);

			// Set Marshal Attributes for return type
			if (!IsCommand && ReturnType.MarshalAsAttribute != null)
			{
					ParameterBuilder pb = methodBuilder.DefineParameter(0, ParameterAttributes.None, null);
					pb.SetCustomAttribute(ReturnType.MarshalAsAttribute);
			}
			
			// ... and the parameters
			for (int i = 1; i <= paramInfos.Length; i++)
			{
				CustomAttributeBuilder b = paramInfos[i-1].MarshalAsAttribute;
				if (b != null)
				{
					ParameterBuilder pb = methodBuilder.DefineParameter(i, ParameterAttributes.None, null);
					pb.SetCustomAttribute(b);
				}
			}

			// Bake the type and get the delegate
			return typeBuilder.CreateType();
		}

		private Delegate CreateMethodDelegate(MethodInfo targetMethod, Type delegateType)
		{
            // Check whether we can skip wrapper
            if (IsExceptionSafe
                && Array.TrueForAll(Parameters, 
                        delegate(XlParameterInfo pi){ return pi.BoxedValueType == null;})
                && (IsCommand || ReturnType.BoxedValueType == null))
            {
                // Create the delegate directly
				return Delegate.CreateDelegate(delegateType, targetMethod);
            }

			// Else we create a dynamic wrapper
			Type[] paramTypes = Array.ConvertAll<XlParameterInfo, Type>(Parameters,
				delegate(XlParameterInfo pi) { return pi.DelegateParamType; });

			DynamicMethod wrapper = new DynamicMethod(
				string.Format("Wrapped_f{0}_{1}", Index, targetMethod.Name),
				IsCommand ? typeof(void) : ReturnType.DelegateParamType, 
				paramTypes, typeof(object), true);
			ILGenerator wrapIL = wrapper.GetILGenerator();
			Label endOfMethod = wrapIL.DefineLabel();

			LocalBuilder retobj = null;
			if (!IsCommand)
			{
				// Make a local to contain the return value
				retobj = wrapIL.DeclareLocal(ReturnType.DelegateParamType);
			}
			if (!IsExceptionSafe)
			{
				// Start the Try block
				wrapIL.BeginExceptionBlock();
			}

			// Generate the body
			// push all the arguments
			for (byte i = 0; i < paramTypes.Length; i++)
			{
                wrapIL.Emit(OpCodes.Ldarg_S, i);
                XlParameterInfo pi = Parameters[i];
                if (pi.BoxedValueType != null)
                {
                    wrapIL.Emit(OpCodes.Unbox_Any , pi.BoxedValueType);
                }
			}
			// Call the real method
			wrapIL.EmitCall(OpCodes.Call, targetMethod, null);
			if (!IsCommand && ReturnType.BoxedValueType != null)
			{
				// Box the return value (which is on the stack)
				wrapIL.Emit(OpCodes.Box, ReturnType.BoxedValueType);
			}
			if (!IsCommand)
			{
				// Store the return value into the local variable
				wrapIL.Emit(OpCodes.Stloc_S, retobj);
			}

			if (!IsExceptionSafe)
			{
				wrapIL.Emit(OpCodes.Leave_S, endOfMethod );
				wrapIL.BeginCatchBlock(typeof(object));
				if (!IsCommand && ReturnType.DelegateParamType == typeof(object))
				{
					// Call Integration.HandleUnhandledException - Exception object is on the stack.
					wrapIL.EmitCall(OpCodes.Call, IntegrationHelpers.UnhandledExceptionHandler, null);
					// Stack now has return value from the ExceptionHandler - Store to local 
					wrapIL.Emit(OpCodes.Stloc_S, retobj);

					//// Create a boxed Excel error value, and set the return object to it
					//wrapIL.Emit(OpCodes.Ldc_I4, IntegrationMarshalHelpers.ExcelError_ExcelErrorValue);
					//wrapIL.Emit(OpCodes.Box, IntegrationMarshalHelpers.GetExcelErrorType());
					//wrapIL.Emit(OpCodes.Stloc_S, retobj);
				}
				else
				{
					// Just ignore the Exception.
					wrapIL.Emit(OpCodes.Pop);
				}
				wrapIL.EndExceptionBlock();
			}
			wrapIL.MarkLabel(endOfMethod);
			if (!IsCommand)
			{
				// Push the return value
				wrapIL.Emit(OpCodes.Ldloc_S, retobj);
			}
			wrapIL.Emit(OpCodes.Ret);
			// End of Wrapper

			return wrapper.CreateDelegate(delegateType);;
		}
    
        // This is the main conversion function called from XlLibrary.RegisterMethods
        public static List<XlMethodInfo> ConvertToXlMethodInfos(List<MethodInfo> methodInfos)
        {
            List<XlMethodInfo> xlMethodInfos = new List<XlMethodInfo>();

            // Set up assembly
            // Examine the methods, built the types and infos
            // Bake the assembly and export the function pointers

            AssemblyBuilder assemblyBuilder;
            ModuleBuilder moduleBuilder;
            assemblyBuilder = AppDomain.CurrentDomain.DefineDynamicAssembly(
                new AssemblyName("ExcelDna.DynamicDelegateAssembly"),
                AssemblyBuilderAccess.Run/*AndSave*/);
            moduleBuilder = assemblyBuilder.DefineDynamicModule("DynamicDelegates");

            foreach  (MethodInfo mi  in methodInfos)
            {
                try
                {
                    XlMethodInfo xlmi = new XlMethodInfo(mi, moduleBuilder);
                    // Add if no Exceptions
                    xlMethodInfos.Add(xlmi);
                }
                catch (DnaMarshalException e)
                {
                    // TODO: What to do here  (maybe logging)?
                    Debug.Print("ExcelDNA -> Inappropriate Method: " + mi.Name + " - " + e.Message);
                }
            }

            //			assemblyBuilder.Save(@"ExcelDna.DynamicDelegateAssembly.dll");
            return xlMethodInfos;
        }
    }

	// TODO: improve information about the problem
	internal class DnaMarshalException : Exception
	{
		public DnaMarshalException(string message) :
			base(message)
		{
		}
	}
}