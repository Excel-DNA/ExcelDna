//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Loader.Logging;

namespace ExcelDna.Loader
{
    internal class XlMethodInfo
    {
        // These three fields are used during construction
        // Either MethodInfo or LambdaExpression must be set, not both
        public MethodInfo MethodInfo;
        public object Target;   // Only used with MethodInfo. Mostly null - can / should we get rid of it? It only lets us use constant objects to invoke against, which is not so useful. Rather allow open delegates?
        public LambdaExpression LambdaExpression;

        // Set and used during contruction and registration
        public GCHandle DelegateHandle; // TODO: What with this - when to clean up? 
        // For cleanup should call DelegateHandle.Free()
        public IntPtr FunctionPointer;
        // TODO: Add Index in JmpTable

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
        public bool SuppressOverwriteError;
        public double RegisterId;   // Set when we register

        public XlParameterInfo[] Parameters;
        public XlParameterInfo ReturnType; // Macro will have ReturnType null (as will native async functions)

        // THROWS: Throws a DnaMarshalException if the method cannot be turned into an XlMethodInfo
        // TODO: Manage errors if things go wrong
        XlMethodInfo(MethodInfo targetMethod, object target, LambdaExpression lambdaExpression, object methodAttribute, List<object> argumentAttributes)
        {
            if ((targetMethod == null && lambdaExpression == null) ||
                (targetMethod != null && lambdaExpression != null) ||
                (target != null && targetMethod == null))
            {
                throw new ArgumentException("Invalid arguments for XlMarshalInfo");
            }

            MethodInfo = targetMethod;
            Target = target;
            LambdaExpression = lambdaExpression;

            // Default Name, Description and Category
            Name = targetMethod?.Name ?? lambdaExpression.Name;
            Description = "";
            Category = ExcelIntegration.DnaLibraryGetName();
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
            MenuName = Category;
            MenuText = null; // Menu is only 

            // Set default IsCommand - overridden by having an [ExcelCommand] or [ExcelFunction] attribute,
            // or by being a native async function.
            // (Must be done before SetAttributeInfo)
            var returnType = targetMethod?.ReturnType ?? lambdaExpression.ReturnType;
            IsCommand = (returnType == typeof(void));

            SetAttributeInfo(methodAttribute);
            // We shortcut the rest of the registration
            if (ExplicitRegistration) return;

            FixHelpTopic();

            // Return type conversion
            // Careful here - native async functions also return void
            if (returnType == typeof(void))
                ReturnType = null;
            else
                ReturnType = new XlParameterInfo(returnType, true, IsExceptionSafe);

            if (targetMethod != null)
            {
                ParameterInfo[] parameters = targetMethod.GetParameters();

                // Parameters - meta-data and type conversion
                Parameters = new XlParameterInfo[parameters.Length];
                for (int i = 0; i < parameters.Length; i++)
                {
                    object argAttrib = null;
                    if (argumentAttributes != null && i < argumentAttributes.Count)
                        argAttrib = argumentAttributes[i];

                    var param = parameters[i];
                    Parameters[i] = new XlParameterInfo(param.Name, param.ParameterType, argAttrib);
                }
            }
            else
            {
                var parameters = lambdaExpression.Parameters;
                // Parameters - meta-data and type conversion
                Parameters = new XlParameterInfo[parameters.Count];
                for (int i = 0; i < parameters.Count; i++)
                {
                    object argAttrib = null;
                    if (argumentAttributes != null && i < argumentAttributes.Count)
                        argAttrib = argumentAttributes[i];

                    var param = parameters[i];
                    Parameters[i] = new XlParameterInfo(param.Name, param.Type, argAttrib);
                }
            }

            // A native async function might still be marked as a command - check and fix.
            // (these have the ExcelAsyncHandle as last parameter)
            // (This check needs the Parameters array to be set up already.)
            if (IsExcelAsyncFunction)
            {
                // It really is a function, though it might return null
                IsCommand = false;
            }
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
            if (attrib is DescriptionAttribute desc)
            {
                Description = desc.Description;
                return;
            }

            if (attrib is ExcelFunctionAttribute func)
            {
                if (func.Name != null)
                    Name = func.Name;
                if (func.Description != null)
                    Description = func.Description;
                if (func.Category != null)
                    Category = func.Category;
                if (func.HelpTopic != null)
                    HelpTopic = func.HelpTopic;
                IsVolatile = func.IsVolatile;
                IsExceptionSafe = func.IsExceptionSafe;
                IsMacroType = func.IsMacroType;
                IsHidden = func.IsHidden;
                IsThreadSafe = (!func.IsMacroType && func.IsThreadSafe);
                // DOCUMENT: IsClusterSafe function MUST NOT be marked as IsMacroType=true and MAY be marked as IsThreadSafe = true.
                //           [xlfRegister (Form 1) page in the Microsoft Excel 2010 XLL SDK Documentation]
                IsClusterSafe = (!func.IsMacroType && func.IsClusterSafe);
                ExplicitRegistration = func.ExplicitRegistration;
                SuppressOverwriteError = func.SuppressOverwriteError;
                IsCommand = false;
            }
            else if (attrib is ExcelCommandAttribute cmd)
            {
                if (cmd.Name != null)
                    Name = cmd.Name;
                if (cmd.Description != null)
                    Description = cmd.Description;
                if (cmd.HelpTopic != null)
                    HelpTopic = cmd.HelpTopic;
                if (cmd.ShortCut != null)
                    ShortCut = cmd.ShortCut;
                if (cmd.MenuName != null)
                    MenuName = cmd.MenuName;
                if (cmd.MenuText != null)
                    MenuText = cmd.MenuText;
//                    IsHidden = isHidden;  // Only for functions.
                IsExceptionSafe = cmd.IsExceptionSafe;
                ExplicitRegistration = cmd.ExplicitRegistration;
                SuppressOverwriteError = cmd.SuppressOverwriteError;

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

        // This is the main conversion function called from XlLibrary.RegisterMethods
        // targets may be null - the typical case
        // Either methods or lambdaExpressions may be null
        // If they are both supplied, then the corresponding entries in the lists must be null in one of the lists
        public static List<XlMethodInfo> ConvertToXlMethodInfos(List<MethodInfo> methods, List<object> targets, List<LambdaExpression> lambdaExpressions, List<object> methodAttributes, List<List<object>> argumentAttributes)
        {
            List<XlMethodInfo> xlMethodInfos = new List<XlMethodInfo>();
            var count = methods?.Count ?? lambdaExpressions.Count;

            for (int i = 0; i < count; i++)
            {
                MethodInfo mi = methods?[i]; // List might be null
                object target = targets?[i]; // List might be null
                LambdaExpression lambda = lambdaExpressions?[i]; // List might be null
                object methodAttrib = (methodAttributes != null && i < methodAttributes.Count) ? methodAttributes[i] : null;
                List<object> argAttribs = (argumentAttributes != null && i < argumentAttributes.Count) ? argumentAttributes[i] : null;
                try
                {
                    XlMethodInfo xlmi = new XlMethodInfo(mi, target, lambda, methodAttrib, argAttribs);
                    // Skip if suppressed
                    if (xlmi.ExplicitRegistration)
                    {
                        Logger.Registration.Info("Suppressing due to ExplictRegistration attribute: '{0}.{1}'", mi?.DeclaringType.Name ?? "<lambda>", mi?.Name ?? lambda?.Name);
                        continue;
                    }
                    xlMethodInfos.Add(xlmi);
                }
                catch (DnaMarshalException e)
                {
                    Logger.Registration.Error(e, "Method not registered due to unsupported signature: '{0}.{1}'", mi?.DeclaringType.Name ?? "<lambda>", mi?.Name ?? lambda?.Name);
                }
            }

            Parallel.ForEach(xlMethodInfos, xlmi => XlDirectMarshal.SetDelegateAndFunctionPointer(xlmi));

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
                    if (att is ExcelFunctionAttribute || att is ExcelCommandAttribute)
                    {
                        // Set last value to this attribute
                        methodAttributes[methodAttributes.Count - 1] = att;
                        break;
                    }
                    if (att is DescriptionAttribute)
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
                        if (att is ExcelArgumentAttribute)
                        {
                            // Set last value to this attribute
                            argAttribs[argAttribs.Count - 1] = att;
                            break;
                        }
                        if (att is DescriptionAttribute)
                        {
                            // Some compatibility - use Description if no ExcelArgument attribute
                            if (argAttribs[argAttribs.Count - 1] == null)
                                argAttribs[argAttribs.Count - 1] = att;
                        }
                    }

                }
            }
        }

        public bool HasReturnType => ReturnType != null;
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
