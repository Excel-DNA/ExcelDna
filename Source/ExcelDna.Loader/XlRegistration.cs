/*
  Copyright (C) 2005-2013 Govert van Drimmelen

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

namespace ExcelDna.Loader
{
    // TODO: Migrate registration to ExcelDna.Integration
    public static class XlRegistration
    {
        static readonly List<XlMethodInfo> registeredMethods = new List<XlMethodInfo>();
        static readonly List<string> addedShortCuts = new List<string>();
        
        // This list is just to give access to the function information for UI enhancement.
        // We should probably replace it with a general extension API when I am comfortable about how that would look.
        // Populated by RecordFunctionInfo(...) below
        // Each entry has: 
        // name, category, helpTopic, argumentNames, [description], [argumentDescription_1] ... [argumentDescription_n].
        static readonly List<List<string>> functionInfo = new List<List<string>>();
        
        public static void RegisterMethods(List<MethodInfo> methods)
        {
            List<object> methodAttributes;
            List<List<object>> argumentAttributes;
            XlMethodInfo.GetMethodAttributes(methods, out methodAttributes, out argumentAttributes);
            RegisterMethodsWithAttributes(methods, methodAttributes, argumentAttributes);
        }

        public static void RegisterMethodsWithAttributes(List<MethodInfo> methods, List<object> methodAttributes, List<List<object>> argumentAttributes)
        {
            Register(methods, null,  methodAttributes, argumentAttributes);
        }

        public static void RegisterDelegatesWithAttributes(List<Delegate> delegates, List<object> methodAttributes, List<List<object>> argumentAttributes)
        {
            // I'm missing LINQ ...
            List<MethodInfo> methods = new List<MethodInfo>();
            List<object> targets = new List<object>();
            for (int i = 0; i < delegates.Count; i++)
            {
                Delegate del = delegates[i];
                // Using del.Method and del.Target from here is a problem 
                // - then we have to deal with the open/closed situation very carefully.
                // We'll pass and invoke the actual delegate, which means the method signature is correct.
                // Overhead should be negligible.
                methods.Add(del.GetType().GetMethod("Invoke"));
                targets.Add(del);
            }
            Register(methods, targets, methodAttributes, argumentAttributes);
        }

        public static List<List<string>> GetFunctionRegistrationInfo()
        {
            return functionInfo;
        }

        static void Register(List<MethodInfo> methods, List<object> targets, List<object> methodAttributes, List<List<object>> argumentAttributes)
        {
            Debug.Assert(targets == null || targets.Count == methods.Count);

            List<XlMethodInfo> xlMethods = XlMethodInfo.ConvertToXlMethodInfos(methods, targets, methodAttributes, argumentAttributes);
            xlMethods.ForEach(RegisterXlMethod);
        }

        private static void RegisterXlMethod(XlMethodInfo mi)
        {
            int index = registeredMethods.Count;
            XlAddIn.SetJump(index, mi.FunctionPointer);
            String exportedProcName = String.Format("f{0}", index);

            object[] registerParameters = GetRegisterParameters(mi, exportedProcName);
            if (!mi.IsCommand && !mi.IsHidden)
            {
                RecordFunctionInfo(registerParameters);
            }

            // Basically suppress problems here !?
            try
            {
                object xlCallResult;
                XlCallImpl.TryExcelImpl(XlCallImpl.xlfRegister, out xlCallResult, registerParameters);
                Debug.Print("Register - XllPath={0}, ProcName={1}, FunctionType={2}, MethodName={3} - Result={4}",
                            registerParameters[0], registerParameters[1], registerParameters[2], registerParameters[3],
                            xlCallResult);
                if (xlCallResult is double)
                {
                    mi.RegisterId = (double) xlCallResult;
                    registeredMethods.Add(mi);
                    if (mi.IsCommand)
                    {
                        RegisterMenu(mi);
                        RegisterShortCut(mi);
                    }
                }
                else
                {
                    // TODO: What to do here? LogDisplay??
                    Debug.Print("Registration Error! - Register call failed for method {0}", mi.Name);
                }
            }
            catch (Exception e)
            {
                // TODO: What to do here? LogDisplay??
                Debug.WriteLine("Registration Error! - " + e.Message);
            }
        }

        // NOTE: We are not currently removing the functions from the Jmp array
        internal static void UnregisterMethods()
        {
            object xlCallResult;

            // Remove menus and ShortCuts
            IntegrationHelpers.RemoveCommandMenus();
            UnregisterShortCuts();

            // Now take out the methods
            foreach (XlMethodInfo mi in registeredMethods)
            {
                // Clear the name and unregister
                XlCallImpl.TryExcelImpl(XlCallImpl.xlfSetName, out xlCallResult, mi.Name);
                XlCallImpl.TryExcelImpl(XlCallImpl.xlfUnregister, out xlCallResult, mi.RegisterId);

                if (!mi.IsCommand)
                {
                    // I follow the advice from X-Cell website to get function out of Wizard (with fix from kh)
                    // clear the new name, and unregister
                    XlCallImpl.TryExcelImpl(XlCallImpl.xlfRegister, out xlCallResult, XlAddIn.PathXll, "xlAutoRemove", "I", mi.Name, IntegrationMarshalHelpers.GetExcelMissingValue(), 2);
                    if (xlCallResult is double)
                    {
                        double fakeRegisterId = (double)xlCallResult;
                        XlCallImpl.TryExcelImpl(XlCallImpl.xlfSetName, out xlCallResult, mi.Name);
                        XlCallImpl.TryExcelImpl(XlCallImpl.xlfUnregister, out xlCallResult, fakeRegisterId);
                    }
                }
            }
            registeredMethods.Clear();
        }

        private static void RegisterMenu(XlMethodInfo mi)
        {
            if (!string.IsNullOrEmpty(mi.MenuName) &&
                !string.IsNullOrEmpty(mi.MenuText))
            {
                IntegrationHelpers.AddCommandMenu(mi.Name, mi.MenuName, mi.MenuText, mi.Description, mi.ShortCut, mi.HelpTopic);
            }
        }

        private static void RegisterShortCut(XlMethodInfo mi)
        {
            if (!string.IsNullOrEmpty(mi.ShortCut))
            {
                object xlCallResult;
                XlCallImpl.TryExcelImpl(XlCallImpl.xlcOnKey, out xlCallResult, mi.ShortCut, mi.Name);
                // CONSIDER: We ignore result and suppress errors - maybe log?
                addedShortCuts.Add(mi.ShortCut);
            }
        }

        private static void UnregisterShortCuts()
        {
            foreach (string shortCut in addedShortCuts)
            {
                // xlcOnKey with no macro name:
                // "If macro_text is omitted, key_text reverts to its normal meaning in Microsoft Excel, 
                // and any special key assignments made with previous ON.KEY functions are cleared."
                object xlCallResult;
                XlCallImpl.TryExcelImpl(XlCallImpl.xlcOnKey, out xlCallResult, shortCut);
            }
        }

        private static object[] GetRegisterParameters(XlMethodInfo mi, string exportedProcName)
        {
            string functionType;
            if (mi.ReturnType != null)
            {
                functionType = mi.ReturnType.XlType;
            }
            else
            {
                if (mi.Parameters.Length == 0)
                {
                    functionType = "";  // OK since no other types will be added
                }
                else
                {
                    // This case is also be used for native async functions
                    functionType = ">"; // Use the void / inplace indicator if needed.
                }
            }

            // TODO: The argument names and descriptions allow some undocumented ",..." form to support paramarray style functions.
            //       E.g. check the FuncSum function in Generic.c in the SDK.
            //       We should try some support for this...
            string argumentNames = "";
            bool showDescriptions = false;
            // For async functions, we need to leave off the last argument
            int numArgumentDescriptions = mi.IsExcelAsyncFunction ? mi.Parameters.Length - 1 : mi.Parameters.Length;
            string[] argumentDescriptions = new string[numArgumentDescriptions];

            for (int j = 0; j < numArgumentDescriptions; j++)
            {
                XlParameterInfo pi = mi.Parameters[j];

                functionType += pi.XlType;

                if (j > 0)
                    argumentNames += ",";
                argumentNames += pi.Name;
                argumentDescriptions[j] = pi.Description;

                if (pi.Description != "")
                    showDescriptions = true;

                // DOCUMENT: Truncate the argument description if it exceeds the Excel limit of 255 characters
                if (j < mi.Parameters.Length - 1)
                {
                    if (!string.IsNullOrEmpty(argumentDescriptions[j]) &&
                        argumentDescriptions[j].Length > 255)
                    {
                        argumentDescriptions[j] = argumentDescriptions[j].Substring(0, 255);
                        Debug.Print("Truncated argument description of {0} in method {1} as Excel limit was exceeded",
                                    pi.Name, mi.Name);
                    }
                }
                else
                {
                    // Last argument - need to deal with extra ". "
                    if (!string.IsNullOrEmpty(argumentDescriptions[j]))
                    {
                        if (argumentDescriptions[j].Length > 253)
                        {
                            argumentDescriptions[j] = argumentDescriptions[j].Substring(0, 253);
                            Debug.Print("Truncated field description of {0} in method {1} as Excel limit was exceeded",
                                        pi.Name, mi.Name);
                        }

                        // DOCUMENT: Here is the patch for the Excel Function Description bug.
                        // DOCUMENT: I add ". " to the last parameter.
                        argumentDescriptions[j] += ". ";
                    }
                }
            } // for each parameter

            // Add async handle
            if (mi.IsExcelAsyncFunction)
                functionType += "X"; // mi.Parameters[mi.Parameters.Length - 1].XlType should be "X" anyway

            // Native async functions cannot be cluster safe
            if (mi.IsClusterSafe && ProcessHelper.SupportsClusterSafe && !mi.IsExcelAsyncFunction)
                functionType += "&";

            if (mi.IsMacroType)
                functionType += "#";

            if (!mi.IsMacroType && mi.IsThreadSafe && XlAddIn.XlCallVersion >= 12)
                functionType += "$";

            if (mi.IsVolatile)
                functionType += "!";
            // DOCUMENT: If # is set and there is an R argument, Excel considers the function volatile anyway.
            // You can call xlfVolatile, false in beginning of function to clear.

            string functionDescription = mi.Description;
            // DOCUMENT: Truncate Description to 253 characters (for all versions)
            functionDescription = Truncate(functionDescription, 253);

            // DOCUMENT: Here is the patch for the Excel Function Description bug.
            // DOCUMENT: I add ". " if the function takes no parameters and has a description.
            if (mi.Parameters.Length == 0 && functionDescription != "")
                functionDescription += ". ";

            // DOCUMENT: When there is no description, we don't add any.
            // This allows the user to work around the Excel bug where an extra parameter is displayed if
            // the function has no parameter but displays a description
            if (mi.Description != "")
                showDescriptions = true;

            int numRegisterParameters;
            // DOCUMENT: Maximum 20 Argument Descriptions when registering using Excel4 function.
            int maxDescriptions = (XlAddIn.XlCallVersion < 12) ? 20 : 245;
            if (showDescriptions)
            {
                numArgumentDescriptions = Math.Min(numArgumentDescriptions, maxDescriptions);
                numRegisterParameters = 10 + numArgumentDescriptions;    // function description + arg descriptions
            }
            else
            {
                // Won't be showing any descriptions.
                numArgumentDescriptions = 0;
                numRegisterParameters = 9;
            }

            // DOCUMENT: Additional truncations of registration info - registration fails with strings longer than 255 chars.
            argumentNames = Truncate(argumentNames, 255);
            string category = Truncate(mi.Category, 255);
            string name = Truncate(mi.Name, 255);
            string helpTopic = (mi.HelpTopic == null || mi.HelpTopic.Length <= 255) ? mi.HelpTopic : "";

            object[] registerParameters = new object[numRegisterParameters];
            registerParameters[0] = XlAddIn.PathXll;
            registerParameters[1] = exportedProcName;
            registerParameters[2] = functionType;
            registerParameters[3] = name;
            registerParameters[4] = argumentNames;
            registerParameters[5] = mi.IsCommand ? 2 /*macro*/
                                                 : (mi.IsHidden ? 0 : 1); /*function*/
            registerParameters[6] = category;
            registerParameters[7] = mi.ShortCut; /*shortcut_text*/
            registerParameters[8] = helpTopic; /*help_topic*/

            if (showDescriptions)
            {
                registerParameters[9] = functionDescription;

                for (int k = 0; k < numArgumentDescriptions; k++)
                {
                    registerParameters[10 + k] = argumentDescriptions[k];
                }
            }

            return registerParameters;
        }

        private static void RecordFunctionInfo(object[] registerParameters)
        {
            // name, category, helpTopic, argumentNames, [description], [argumentDescription_1] ... [argumentDescription_n].
            List<string> info = new List<string>();
            info.Add((string)registerParameters[3]);    // name
            info.Add((string)registerParameters[6]);    // category
            info.Add((string)registerParameters[8]);    // helpTopic
            info.Add((string)registerParameters[4]);    // argumentNames
            if (registerParameters.Length >= 10)
            {
                info.Add((string)registerParameters[9]); // Description
            }
            for (int k = 10; k < registerParameters.Length; k++)
            {
                info.Add((string)registerParameters[k]);  // argumentDescription
            }
            functionInfo.Add(info);
        }

        static string Truncate(string s, int length)
        {
            if (s == null || s.Length <= length) return s;
            return s.Substring(0, length);
        }
    }
}
