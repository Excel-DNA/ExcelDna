using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;

namespace ExcelDna.Loader
{
    // TODO: Migrate registration to ExcelDna.Integration
    public static class XlRegistration
    {
        static readonly List<XlMethodInfo> registeredMethods = new List<XlMethodInfo>();
        static List<string> addedShortCuts = new List<string>();

        public static void RegisterMethods(List<MethodInfo> methods)
        {
            List<object> methodAttributes;
            List<List<object>> argumentAttributes;
            XlMethodInfo.GetMethodAttributes(methods, out methodAttributes, out argumentAttributes);
            RegisterMethodsWithAttributes(methods, methodAttributes, argumentAttributes);
        }

        public static void RegisterMethodsWithAttributes(List<MethodInfo> methods, List<object> methodAttributes, List<List<object>> argumentAttributes)
        {
            List<XlMethodInfo> xlMethods = XlMethodInfo.ConvertToXlMethodInfos(methods, methodAttributes, argumentAttributes);
            xlMethods.ForEach(RegisterXlMethod);
        }

        private static void RegisterXlMethod(XlMethodInfo mi)
        {
            int index = registeredMethods.Count;
            XlAddIn.SetJump(index, mi.FunctionPointer);
            String exportedProcName = String.Format("f{0}", index);

            object[] registerParameters = GetRegisterParameters(mi, exportedProcName);

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

            // DOCUMENT: There is a bug? in Excel 2007 that limits the total argumentname string to 255 chars.
            // TODO: Check whether this is fixed in Excel 2010/2013 yet.
            // DOCUMENT: I truncate the argument string for all versions.
            if (argumentNames.Length > 255)
                argumentNames = argumentNames.Substring(0, 255);

            // DOCUMENT: Here is the patch for the Excel Function Description bug.
            // DOCUMENT: I add ". " if the function takes no parameters and has a description.
            string functionDescription = mi.Description;
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

            object[] registerParameters = new object[numRegisterParameters];
            registerParameters[0] = XlAddIn.PathXll;
            registerParameters[1] = exportedProcName;
            registerParameters[2] = functionType;
            registerParameters[3] = mi.Name;
            registerParameters[4] = argumentNames;
            registerParameters[5] = mi.IsCommand ? 2 /*macro*/
                                                 : (mi.IsHidden ? 0 : 1); /*function*/
            registerParameters[6] = mi.Category;
            registerParameters[7] = mi.ShortCut; /*shortcut_text*/
            registerParameters[8] = mi.HelpTopic; /*help_topic*/ ;

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
    }
}
