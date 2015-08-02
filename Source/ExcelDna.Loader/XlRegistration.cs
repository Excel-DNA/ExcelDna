//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using ExcelDna.Loader.Logging;

namespace ExcelDna.Loader
{
    // TODO: Migrate registration to ExcelDna.Integration
    public static class XlRegistration
    {
        static readonly List<XlMethodInfo> registeredMethods = new List<XlMethodInfo>();
        static readonly List<string> addedShortCuts = new List<string>();
        
        // This list is just to give access to the registration details for UI enhancement.
        // Each entry corresponds exactly to the xlfRegister call (except first entry with xllPath is cleared) 
        // - max length of each array is 255.
        static readonly List<object[]> registrationInfo = new List<object[]>();
        static double registrationInfoVersion = 0.0; // Incremented every time the registration changes, used by GetRegistrationInfo to short-circuit.
        
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

        // This function provides access to the registration info from an IntelliSense provider.
        // To allow polling, we return as the first row a (double) version which can be passed to short-circuit the call if nothing has changed.
        // The signature and behaviour should be flexible enough to allow future non-breaking extension.
        public static object GetRegistrationInfo(object param)
        {
            if (param is double && (double)param == registrationInfoVersion)
            {
                // Short circuit, to prevent returning the whole string story every time, allowing fast polling.
                return null;
            }

            // Copy from the jagged List to a 2D array with 255 columns
            // (missing bits are returned as null, which is marshaled to XlEmpty)
            object[,] result = new object[registrationInfo.Count + 1, 255];
            // Return xll path and registrationVersion in first row
            result[0, 0] = XlAddIn.PathXll;
            result[0, 1] = registrationInfoVersion;

            // Other rows contain the registation info 
            for (int i = 0; i < registrationInfo.Count; i++)
            {
                int resultRow = i + 1;
                object[] info = registrationInfo[i];
                for (int j = 0; j < 255; j++)
                {
                    if (j >= info.Length)
                    {
                        // Done with this row
                        break;
                    }
                    result[resultRow, j] = info[j];
                }
            }
            return result;
        }

        static void Register(List<MethodInfo> methods, List<object> targets, List<object> methodAttributes, List<List<object>> argumentAttributes)
        {
            Debug.Assert(targets == null || targets.Count == methods.Count);
            Logger.Registration.Verbose("Registering {0} methods", methods.Count);
            List<XlMethodInfo> xlMethods = XlMethodInfo.ConvertToXlMethodInfos(methods, targets, methodAttributes, argumentAttributes);
            xlMethods.ForEach(RegisterXlMethod);
            // Increment the registration version (safe to call a few times)
            registrationInfoVersion += 1.0;
        }

        static void RegisterXlMethod(XlMethodInfo mi)
        {
            int index = registeredMethods.Count;
            XlAddIn.SetJump(index, mi.FunctionPointer);
            String exportedProcName = String.Format("f{0}", index);

            object[] registerParameters = GetRegisterParameters(mi, exportedProcName);

            if (registrationInfo.Exists(ri => ((string)ri[3]).Equals((string)registerParameters[3], StringComparison.OrdinalIgnoreCase)))
            {
                // This function will be registered with a name that has already been used (by this add-in)
                // This logged as an error, but the registration continues - the last function with the name wins, for backward compatibility.
                Logger.Registration.Error("Repeated function name: '{0}' - previous registration will be overwritten. ", registerParameters[3]);
            }
            
            // Basically suppress problems here !?
            try
            {
                object xlCallResult;
                XlCallImpl.TryExcelImpl(XlCallImpl.xlfRegister, out xlCallResult, registerParameters);
                Logger.Registration.Info("Register - XllPath={0}, ProcName={1}, FunctionType={2}, Name={3} - Result={4}",
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
                    Logger.Registration.Error("xlfRegister call failed for function or command: '{0}'", mi.Name);
                }
                // Now clear out the xll path and store the parameters to support RegistrationInfo access.
                registerParameters[0] = null;
                registrationInfo.Add(registerParameters);
            }
            catch (Exception e)
            {
                Logger.Registration.Error(e, "Registration failed for function or command: '{0}'", mi.Name);
            }
        }

        // NOTE: We are not currently removing the functions from the Jmp array
        //       That would be needed to do a proper per-method deregistration,
        //       together with a garbage-collectable story for the wrapper methods and delegates, 
        //       instead of the currently runtime-compiled and loaded assemblies.
        internal static void UnregisterMethods()
        {
            object xlCallResult;

            // Remove menus and ShortCuts
            IntegrationHelpers.RemoveCommandMenus();
            UnregisterShortCuts();

            // Now take out the methods
            foreach (XlMethodInfo mi in registeredMethods)
            {
                if (mi.IsCommand)
                {
                    // Clear the name and unregister
                    XlCallImpl.TryExcelImpl(XlCallImpl.xlfSetName, out xlCallResult, mi.Name);
                    XlCallImpl.TryExcelImpl(XlCallImpl.xlfUnregister, out xlCallResult, mi.RegisterId);
                }
                else
                {
                    // And Unregister the real function
                    XlCallImpl.TryExcelImpl(XlCallImpl.xlfUnregister, out xlCallResult, mi.RegisterId);
                    // I follow the advice from X-Cell website to get function out of Wizard (with fix from kh)
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
            registrationInfo.Clear();
        }

        static void RegisterMenu(XlMethodInfo mi)
        {
            if (!string.IsNullOrEmpty(mi.MenuName) &&
                !string.IsNullOrEmpty(mi.MenuText))
            {
                IntegrationHelpers.AddCommandMenu(mi.Name, mi.MenuName, mi.MenuText, mi.Description, mi.ShortCut, mi.HelpTopic);
            }
        }

        static void RegisterShortCut(XlMethodInfo mi)
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
                    argumentNames += ",";   // TODO: Should this be a comma, or the Excel list separator?
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
                        Logger.Registration.Warn("Truncated argument description of '{0}' in function '{1}'", pi.Name, mi.Name);
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
                            Logger.Registration.Warn("Truncated final argument description of function '{0}'", mi.Name);
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

            // DOCUMENT: Truncate Description to 253 characters (for all versions)
            string functionDescription = Truncate(mi.Description, 253, "function description", mi);

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
            argumentNames = Truncate(argumentNames, 255, "argument names", mi);
            argumentNames = argumentNames.TrimEnd(','); // Also trim trailing commas (for params case)
            string category = Truncate(mi.Category, 255, "Category name", mi);
            string name = Truncate(mi.Name, 255, "Name", mi);
            string helpTopic = string.Empty;
            if (mi.HelpTopic != null)
            {
                if (mi.HelpTopic.Length > 255)
                {
                    // Can't safely truncate the help link
                    Logger.Registration.Warn("Ignoring HelpTopic of function '{0}' - too long", mi.Name);
                }
                else
                {
                    // It's OK
                    helpTopic = mi.HelpTopic;
                }
            }

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

        static string Truncate(string s, int length, string logDetail, XlMethodInfo mi)
        {
            if (s == null || s.Length <= length) return s;
            Logger.Registration.Warn("Truncated " + logDetail + " of function '{0}'", mi.Name);
            return s.Substring(0, length);
        }
    }
}
