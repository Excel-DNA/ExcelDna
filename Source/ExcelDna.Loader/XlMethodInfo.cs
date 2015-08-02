//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using ExcelDna.Loader.Logging;

namespace ExcelDna.Loader
{
    internal class XlMethodInfo
    {
        public static int Index = 0;

        public GCHandle DelegateHandle; // TODO: What with this - when to clean up? 
        // For cleanup should call DelegateHandle.Free()
        public IntPtr FunctionPointer;

        // Info for Excel Registration
        public bool IsCommand;
        public string Name; // Name of UDF/Macro in Excel
        public string Description;
        public bool IsHidden; // For Functions only
        public string ShortCut; // For macros only
        public string MenuName; // For macros only
        public string MenuText; // For macros only
        public string Category;
        public bool IsVolatile;
        public bool IsExceptionSafe;
        public bool IsMacroType;
        public bool IsThreadSafe; // For Functions only
        public bool IsClusterSafe; // For Functions only
        public string HelpTopic;
        public bool ExplicitRegistration;
        public double RegisterId;

        public XlParameterInfo[] Parameters;
        public XlParameterInfo ReturnType; // Macro will have ReturnType null (as will native async functions)

        // THROWS: Throws a DnaMarshalException if the method cannot be turned into an XlMethodInfo
        // TODO: Manage errors if things go wrong
        XlMethodInfo(ModuleBuilder modBuilder, MethodInfo targetMethod, object target, object methodAttribute, List<object> argumentAttributes)
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
            ExplicitRegistration = false;

            ShortCut = "";
            // DOCUMENT: Default MenuName is the library name
            // but menu is only added if at least the MenuText is set.
            MenuName = IntegrationHelpers.DnaLibraryGetName();
            MenuText = null; // Menu is only 

            // Set default IsCommand - overridden by having an [ExcelCommand] attribute,
            // or by being a native async function.
            // (Must be done before SetAttributeInfo)
            IsCommand = (targetMethod.ReturnType == typeof(void));

            SetAttributeInfo(methodAttribute);
            // We shortcut the rest of the registration
            if (ExplicitRegistration) return;

            FixHelpTopic();

            // Return type conversion
            // Careful here - native async functions also return void
            if (targetMethod.ReturnType == typeof(void))
            {
                ReturnType = null;
            }
            else
            {
                ReturnType = new XlParameterInfo(targetMethod.ReturnType, true, IsExceptionSafe);
            }

            ParameterInfo[] parameters = targetMethod.GetParameters();
            
            // Parameters - meta-data and type conversion
            Parameters = new XlParameterInfo[parameters.Length];
            for (int i = 0; i < parameters.Length; i++)
            {
                object argAttrib = null;
                if ( argumentAttributes != null && i < argumentAttributes.Count)
                    argAttrib = argumentAttributes[i];
                 Parameters[i] = new XlParameterInfo(parameters[i], argAttrib);
            }

            // A native async function might still be marked as a command - check and fix.
            // (these have the ExcelAsyncHandle as last parameter)
            // (This check needs the Parameters array to be set up already.)
            if (IsExcelAsyncFunction)
            {
                // It really is a function, though it might return null
                IsCommand = false;
            }

            // Create the delegate type, wrap the targetMethod and create the delegate

            // CONSIDER: Currently we need a special delegate type here so that we can hook on the marshaling attributes.
            //           Future version might do straight-forward marshaling, so we can get rid of these types (just use generic methods)
            
            // FirstArgument (if received) is not used in the delegate type created ...
            Type delegateType = CreateDelegateType(modBuilder);
            // ... but is baked into the delegate itself.
            Delegate xlDelegate = CreateMethodDelegate(delegateType, targetMethod, target);

            // Need to add a reference to prevent garbage collection of our delegate
            // Don't need to pin, according to 
            // "How to: Marshal Callbacks and Delegates Using C++ Interop"
            // Currently this delegate is never released
            // TODO: Clean up properly
            DelegateHandle = GCHandle.Alloc(xlDelegate);
            FunctionPointer = Marshal.GetFunctionPointerForDelegate(xlDelegate);
        }

        // Native async functions have a final parameter that is an ExcelAsyncHandle.
        public bool IsExcelAsyncFunction 
        { 
            get 
            { 
                return Parameters.Length > 0 && Parameters[Parameters.Length - 1].IsExcelAsyncHandle; 
            } 
        }

        // Basic setup - get description, category etc.
        void SetAttributeInfo(object attrib)
        {
            if (attrib == null) return;

            // DOCUMENT: Description in ExcelFunctionAtribute overrides DescriptionAttribute
            // DOCUMENT: Default Category is Current Library Name.
            // Get System.ComponentModel.DescriptionAttribute
            // Search through attribs for Description
            System.ComponentModel.DescriptionAttribute desc =
                attrib as System.ComponentModel.DescriptionAttribute;
            if (desc != null)
            {
                Description = desc.Description;
                return;
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
            //    IsExceptionSafe = xlcmd.IsExceptionSafe;
            //    IsCommand = true;
            //}

            Type attribType = attrib.GetType();
                
            if (TypeHelper.TypeHasAncestorWithFullName(attribType, "ExcelDna.Integration.ExcelFunctionAttribute"))
            {
                string name = (string) attribType.GetField("Name").GetValue(attrib);
                string description = (string) attribType.GetField("Description").GetValue(attrib);
                string category = (string) attribType.GetField("Category").GetValue(attrib);
                string helpTopic = (string) attribType.GetField("HelpTopic").GetValue(attrib);
                bool isVolatile = (bool) attribType.GetField("IsVolatile").GetValue(attrib);
                bool isExceptionSafe = (bool) attribType.GetField("IsExceptionSafe").GetValue(attrib);
                bool isMacroType = (bool) attribType.GetField("IsMacroType").GetValue(attrib);
                bool isHidden = (bool) attribType.GetField("IsHidden").GetValue(attrib);
                bool isThreadSafe = (bool) attribType.GetField("IsThreadSafe").GetValue(attrib);
                bool isClusterSafe = (bool)attribType.GetField("IsClusterSafe").GetValue(attrib);
                bool explicitRegistration = (bool)attribType.GetField("ExplicitRegistration").GetValue(attrib);
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
                ExplicitRegistration = explicitRegistration;
                IsCommand = false;
            }
            else if (TypeHelper.TypeHasAncestorWithFullName(attribType, "ExcelDna.Integration.ExcelCommandAttribute"))
            {
                string name = (string) attribType.GetField("Name").GetValue(attrib);
                string description = (string) attribType.GetField("Description").GetValue(attrib);
                string helpTopic = (string) attribType.GetField("HelpTopic").GetValue(attrib);
                string shortCut = (string) attribType.GetField("ShortCut").GetValue(attrib);
                string menuName = (string) attribType.GetField("MenuName").GetValue(attrib);
                string menuText = (string) attribType.GetField("MenuText").GetValue(attrib);
//                    bool isHidden = (bool)attribType.GetField("IsHidden").GetValue(attrib);
                bool isExceptionSafe = (bool) attribType.GetField("IsExceptionSafe").GetValue(attrib);
                bool explicitRegistration = (bool)attribType.GetField("ExplicitRegistration").GetValue(attrib);

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
                ExplicitRegistration = explicitRegistration;

                // Override IsCommand, even though this 'macro' might have a return value.
                // Allow for more flexibility in what kind of macros are supported, particularly for calling
                // via Application.Run.
                IsCommand = true;   
            }
        }

        void FixHelpTopic()
        {
            // Make HelpTopic without full path relative to xllPath
            if (string.IsNullOrEmpty(HelpTopic))
            {
                return;
            }
           // DOCUMENT: If HelpTopic is not rooted - it is expanded relative to .xll path.
            // If http url does not end with !0 it is appended.
            // I don't think https is supported, but it should not be considered an 'unrooted' path anyway.
            // I could not get file:/// working (only checked with Excel 2013)
            if (HelpTopic.StartsWith("http://") || HelpTopic.StartsWith("https://") || HelpTopic.StartsWith("file://"))
            {
                if (!HelpTopic.EndsWith("!0"))
                {
                    HelpTopic = HelpTopic + "!0";
                }
            }
            else if (!Path.IsPathRooted(HelpTopic))
            {
                HelpTopic = Path.Combine(Path.GetDirectoryName(XlAddIn.PathXll), HelpTopic);
            }
        }

        // Delegate type created here ignores firstArgument (if we have one)
        Type CreateDelegateType(ModuleBuilder modBuilder)
        {
            TypeBuilder typeBuilder;
            MethodBuilder methodBuilder;
            XlParameterInfo[] paramInfos = Parameters;
            Type[] paramTypes = Array.ConvertAll<XlParameterInfo, Type>(paramInfos,
                                                                        delegate(XlParameterInfo pi)
                                                                            { return pi.DelegateParamType; });

            // Create a delegate that has the same signature as the method we would like to hook up to
            typeBuilder = modBuilder.DefineType("f" + Index++ + "Delegate",
                                                TypeAttributes.Class | TypeAttributes.Public |
                                                TypeAttributes.Sealed,
                                                typeof (System.MulticastDelegate));
            ConstructorBuilder constructorBuilder = typeBuilder.DefineConstructor(
                MethodAttributes.RTSpecialName | MethodAttributes.HideBySig |
                MethodAttributes.Public, CallingConventions.Standard,
                new Type[] {typeof (object), typeof (int)});
            constructorBuilder.SetImplementationFlags(MethodImplAttributes.Runtime |
                                                      MethodImplAttributes.Managed);

            // Build up the delegate
            // Define the Invoke method for the delegate
            methodBuilder = typeBuilder.DefineMethod("Invoke",
                                                     MethodAttributes.Public | MethodAttributes.HideBySig |
                                                     MethodAttributes.NewSlot | MethodAttributes.Virtual,
                                                     HasReturnType ? ReturnType.DelegateParamType : typeof(void),
                                                     // What here for macro? null or Void ?
                                                     paramTypes);
            methodBuilder.SetImplementationFlags(MethodImplAttributes.Runtime |
                                                 MethodImplAttributes.Managed);

            // Set Marshal Attributes for return type
            if (HasReturnType && ReturnType.MarshalAsAttribute != null)
            {
                ParameterBuilder pb = methodBuilder.DefineParameter(0, ParameterAttributes.None, null);
                pb.SetCustomAttribute(ReturnType.MarshalAsAttribute);
            }

            // ... and the parameters
            for (int i = 1; i <= paramInfos.Length; i++)
            {
                CustomAttributeBuilder b = paramInfos[i - 1].MarshalAsAttribute;
                if (b != null)
                {
                    ParameterBuilder pb = methodBuilder.DefineParameter(i, ParameterAttributes.None, null);
                    pb.SetCustomAttribute(b);
                }
            }

            // Bake the type and get the delegate
            return typeBuilder.CreateType();
        }

        Delegate CreateMethodDelegate(Type delegateType, MethodInfo targetMethod, object target)
        {
            bool isInstanceMethod = !targetMethod.IsStatic;
            // We expect static methods to have target null
            Debug.Assert(isInstanceMethod || target == null);

            // Check whether we can skip wrapper completely
            // TODO: Change this - we should always wrap in an exception handler, but for double return values
            //       we let IsExceptionSafe mean we do a 1/0 in the exception handler, which Excel will catch and handle as #NUM!
            //       For other types it shouldn't matter...
            if (IsExceptionSafe
                && Array.TrueForAll(Parameters,
                                    delegate(XlParameterInfo pi) { return pi.BoxedValueType == null; })
                && (!HasReturnType || ReturnType.BoxedValueType == null))
            {
                // Create the delegate directly
                if (isInstanceMethod)
                {
                    // Can't be DynamicMethod
                    return Delegate.CreateDelegate(delegateType, target, targetMethod);
                }
                if (targetMethod is DynamicMethod)
                        return ((DynamicMethod)targetMethod).CreateDelegate(delegateType);
                    return Delegate.CreateDelegate(delegateType, targetMethod);
            }

            // DateTime input parameters are never exception safe - we need to be able to fail out of the 
            // marshaler when the passed-in argument is an invalid date.
            bool emitExceptionHandler = 
                !IsExceptionSafe || 
                Array.Exists(Parameters, 
                             delegate(XlParameterInfo pi) { return pi.BoxedValueType == typeof(DateTime); });

            // Now we create a dynamic wrapper
            Type[] paramTypes = Array.ConvertAll<XlParameterInfo, Type>(Parameters,
                                                                        delegate(XlParameterInfo pi)
                                                                            { return pi.DelegateParamType; });
            if (isInstanceMethod)
            {
                Type[] allParams = new Type[paramTypes.Length + 1];
                allParams[0] = typeof(object);
                Array.Copy(paramTypes, 0, allParams, 1, paramTypes.Length);
                paramTypes = allParams;
            }

            DynamicMethod wrapper = new DynamicMethod(
                string.Format("Wrapped_f{0}_{1}", Index, targetMethod.Name),
                HasReturnType ? ReturnType.DelegateParamType : typeof(void),
                paramTypes, typeof (object), true);
            ILGenerator wrapIL = wrapper.GetILGenerator();
            Label endOfMethod = wrapIL.DefineLabel();

            LocalBuilder retobj = null;
            if (HasReturnType)
            {
                // Make a local to contain the return value
                retobj = wrapIL.DeclareLocal(ReturnType.DelegateParamType);
            }
            if (emitExceptionHandler)
            {
                // Start the Try block
                wrapIL.BeginExceptionBlock();
            }

            // Generate the body - push all the arguments, including the target for instance methods
            if (isInstanceMethod)
            {
                // First is the target of the delegate
                wrapIL.Emit(OpCodes.Ldarg_S, 0);
            }
            for (byte i = 0; i < Parameters.Length; i++)
            {
                if (i < 255)
                {
                    byte argIndex = isInstanceMethod ? (byte)(i + 1) : i;
                    wrapIL.Emit(OpCodes.Ldarg_S, argIndex);
                }
                else
                {
                    short argIndex = isInstanceMethod ? (short)(i + 1) : i;
                    wrapIL.Emit(OpCodes.Ldarg, argIndex);
                }
                XlParameterInfo pi = Parameters[i];
                if (pi.BoxedValueType != null)
                {
                    wrapIL.Emit(OpCodes.Unbox_Any, pi.BoxedValueType);
                }
            }
            // Call the real method
            wrapIL.EmitCall(OpCodes.Call, targetMethod, null);
            if (HasReturnType)
            {
                if (ReturnType.BoxedValueType != null)
                {
                    // Box the return value (which is on the stack)
                    wrapIL.Emit(OpCodes.Box, ReturnType.BoxedValueType);
                }
                // Store the return value into the local variable
                wrapIL.Emit(OpCodes.Stloc_S, retobj);
            }

            if (emitExceptionHandler)
            {
                wrapIL.Emit(OpCodes.Leave_S, endOfMethod);
                wrapIL.BeginCatchBlock(typeof (object));
                if (HasReturnType && ReturnType.DelegateParamType == typeof (object))
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
            if (HasReturnType)
            {
                // Push the return value
                wrapIL.Emit(OpCodes.Ldloc_S, retobj);
            }
            wrapIL.Emit(OpCodes.Ret);
            // End of Wrapper

            if (isInstanceMethod)
            {
                return wrapper.CreateDelegate(delegateType, target);
            }
            return wrapper.CreateDelegate(delegateType);
        }

        // This is the main conversion function called from XlLibrary.RegisterMethods
        // targets may be null - the typical case
        public static List<XlMethodInfo> ConvertToXlMethodInfos(List<MethodInfo> methods, List<object> targets, List<object> methodAttributes, List<List<object>> argumentAttributes)
        {
            List<XlMethodInfo> xlMethodInfos = new List<XlMethodInfo>();

            // Set up assembly
            // Examine the methods, built the types and infos
            // Bake the assembly and export the function pointers

            AssemblyBuilder assemblyBuilder;
            ModuleBuilder moduleBuilder;
            assemblyBuilder = AppDomain.CurrentDomain.DefineDynamicAssembly(
                new AssemblyName("ExcelDna.DynamicDelegateAssembly"),
                AssemblyBuilderAccess.Run /*AndSave*/);
            moduleBuilder = assemblyBuilder.DefineDynamicModule("DynamicDelegates");

            for (int i = 0; i < methods.Count; i++)
            {
                MethodInfo mi  = methods[i];
                object target = (targets == null) ? null : targets[i];
                object methodAttrib = (methodAttributes != null && i < methodAttributes.Count) ? methodAttributes[i] : null;
                List<object> argAttribs = (argumentAttributes != null && i < argumentAttributes.Count) ? argumentAttributes[i] : null;
                try
                {
                    XlMethodInfo xlmi = new XlMethodInfo(moduleBuilder, mi, target, methodAttrib, argAttribs);
                    // Skip if suppressed
                    if (xlmi.ExplicitRegistration)
                    {
                        Logger.Registration.Info("Suppressing due to ExplictRegistration attribute: '{0}.{1}'", mi.DeclaringType.Name, mi.Name);
                        continue;
                    }
                    // otherwise add
                    xlMethodInfos.Add(xlmi);
                }
                catch (DnaMarshalException e)
                {
                    Logger.Registration.Error(e, "Method not registered due to unsupported signature: '{0}.{1}'", mi.DeclaringType.Name, mi.Name);
                }
            }

            //			assemblyBuilder.Save(@"ExcelDna.DynamicDelegateAssembly.dll");
            return xlMethodInfos;
        }

        public static void GetMethodAttributes(List<MethodInfo> methodInfos, out List<object> methodAttributes, out List<List<object>> argumentAttributes)
        {
            methodAttributes = new List<object>();
            argumentAttributes = new List<List<object>>();
            foreach (MethodInfo method in methodInfos)
            {
                // If we don't find an attribute, we'll set a null in the list at a token
                methodAttributes.Add(null);
                foreach (object att in method.GetCustomAttributes(false))
                {
                    Type attType = att.GetType();
                    if (TypeHelper.TypeHasAncestorWithFullName(attType, "ExcelDna.Integration.ExcelFunctionAttribute") ||
                        TypeHelper.TypeHasAncestorWithFullName(attType, "ExcelDna.Integration.ExcelCommandAttribute"))
                    {
                        // Set last value to this attribute
                        methodAttributes[methodAttributes.Count - 1] = att;
                        break;
                    }
                    if (att is System.ComponentModel.DescriptionAttribute)
                    {
                        // Some compatibility - use Description if no Excel* attribute
                        if (methodAttributes[methodAttributes.Count - 1] == null)
                            methodAttributes[methodAttributes.Count - 1] = att;
                    }
                }

                List<object> argAttribs = new List<object>();
                argumentAttributes.Add(argAttribs);

                foreach (ParameterInfo param in method.GetParameters())
                {
                    // If we don't find an attribute, we'll set a null in the list at a token
                    argAttribs.Add(null);
                    foreach (object att in param.GetCustomAttributes(false))
                    {
                        Type attType = att.GetType();
                        if (TypeHelper.TypeHasAncestorWithFullName(attType, "ExcelDna.Integration.ExcelArgumentAttribute"))
                        {
                            // Set last value to this attribute
                            argAttribs[argAttribs.Count - 1] = att;
                            break;
                        }
                        if (att is System.ComponentModel.DescriptionAttribute)
                        {
                            // Some compatibility - use Description if no ExcelArgument attribute
                            if (argAttribs[argAttribs.Count - 1] == null)
                                argAttribs[argAttribs.Count - 1] = att;
                        }
                    }

                }
            }
        }

        bool HasReturnType { get { return ReturnType != null; } }

    }

    internal static class TypeHelper
    {   
        internal static bool TypeHasAncestorWithFullName(Type type, string fullName)
        {
            if (type == null) return false;
            if (type.FullName == fullName) return true;
            return TypeHasAncestorWithFullName(type.BaseType, fullName);
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