//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using ExcelDna.Logging;

namespace ExcelDna.Integration
{
    static class RegistrationInfo
    {
        static string _registrationInfoName = null;
        static double _registrationInfoRegistrationId;
        // We register the two extension helpers:
        // * SyncMacro_XXXXX which is used for the thread-safe RTD support and QueueAsMacro mechanism
        // * RegistrationInfo_XXXX which is used to retrieve the xlfRegister call info.
        // In both cases the XXXX is replaced with the Guid from the .xll path
        internal static void Register()
        {
            // RegistrationInfo is supported for Excel 2007+ only
            if (ExcelDnaUtil.ExcelVersion < 12.0) return;

            _registrationInfoName = RegistrationInfoName(ExcelDnaUtil.XllPath);

            object[] registerParameters = new object[6];
            registerParameters[0] = ExcelDnaUtil.XllPath;
            registerParameters[1] = "RegistrationInfo";
            registerParameters[2] = "QQ"; // Takes XLOPER12, returns XLOPER12
            registerParameters[3] = _registrationInfoName;
            registerParameters[4] = null;
            registerParameters[5] = 0; // hidden function

            object xlCallResult;
            XlCall.TryExcel(XlCall.xlfRegister, out xlCallResult, registerParameters);
            Logger.Registration.Verbose("Register RegistrationInfo - XllPath={0}, ProcName={1}, FunctionType={2}, MethodName={3} - Result={4}",
                registerParameters[0], registerParameters[1], registerParameters[2], registerParameters[3], xlCallResult);
            if (xlCallResult is double)
            {
                _registrationInfoRegistrationId = (double)xlCallResult;
            }
            else
            {
                throw new InvalidOperationException("Synchronization macro registration failed.");
            }
        }

        internal static void Unregister()
        {
            object xlCallResult;
            XlCall.TryExcel(XlCall.xlfUnregister, out xlCallResult, _registrationInfoRegistrationId);
            XlCall.TryExcel(XlCall.xlfRegister, out xlCallResult, ExcelDnaUtil.XllPath, "xlAutoRemove", "I", _registrationInfoName, ExcelMissing.Value, 2);
            if (xlCallResult is double)
            {
                double fakeRegisterId = (double)xlCallResult;
                XlCall.TryExcel(XlCall.xlfSetName, out xlCallResult, _registrationInfoName);
                XlCall.TryExcel(XlCall.xlfUnregister, out xlCallResult, fakeRegisterId);
            }
        }

        // Name of the Registration Helper function, registered in ExcelIntegration.RegisterRegistrationInfo
        static string RegistrationInfoName(string xllPath)
        {
            return "RegistrationInfo_" + ExcelDnaUtil.GuidFromXllPath(xllPath).ToString("N");
        }

        // Public function to get registration info for this or other Excel-DNA .xlls
        // Return #N/A if the matching RegistrationInfo function is not found.
        internal static object GetRegistrationInfo(string xllPath, double registrationInfoVersion)
        {
            string regInfoName = RegistrationInfoName(xllPath);
            object result;

            // To prevent the error dialog, we first call xlfEvaluate with the name
            if (XlCall.Excel(XlCall.xlfEvaluate, regInfoName) is double &&
                XlCall.TryExcel(XlCall.xlUDF, out result, regInfoName, registrationInfoVersion) == XlCall.XlReturn.XlReturnSuccess)
            {
                    return result;
            }

            return ExcelError.ExcelErrorNA;
        }
    }
}
