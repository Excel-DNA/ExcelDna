#if COM_GENERATED

using ExcelDna.Integration.ComInterop.Generator.Interfaces;
using ExcelDna.Integration.Extensibility;
using ExcelDna.Integration.Win32;
using System;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices.Marshalling;

namespace ExcelDna.Integration.ComInterop.Generator
{
    [GeneratedComClass]
    internal partial class ExcelRibbon : ExcelComAddIn, Interfaces.IDTExtensibility2, Interfaces.IRibbonExtensibility
    {
        private MethodInfo[] methods;
        private CustomUI.IExcelRibbon customRibbon;

        public ExcelRibbon(ITypeHelper t)
        {
            methods = t.Methods.ToArray();
            this.customRibbon = t.CreateInstance() as CustomUI.IExcelRibbon;
        }

        public int GetTypeInfoCount(out uint pctinfo)
        {
            throw new NotImplementedException();
        }

        public int GetTypeInfo(uint iTInfo, uint lcid, out nint ppTInfo)
        {
            throw new NotImplementedException();
        }

        public int GetIDsOfNames(Guid riid, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 2)] string[] rgszNames, uint cNames, uint lcid, [In][Out][MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)] int[] rgDispId)
        {
            for (int i = 0; i < cNames; ++i)
                rgDispId[i] = (rgszNames[i] == "LoadImage") ? methods.Length : Array.FindIndex(methods, m => m?.Name == rgszNames[i]);

            return 0;
        }

        public int Invoke(int dispIdMember, Guid riid, uint lcid, INVOKEKIND wFlags, [MarshalUsing(typeof(Generator.Interfaces.DispParamsMarshaller))] in Generator.Interfaces.DispParams pDispParams, nint pVarResult, nint pExcepInfo, nint puArgErr)
        {
            try
            {
                if (dispIdMember >= 0 && dispIdMember < methods.Length)
                {
                    // Ribbon callback - e.g. onAction / onLoad / getLabel / getVisible / getEnabled / onChange / getItemLabel ...
                    // We bind each Excel-supplied argument to the matching callback parameter, then marshal the return value (if any) back to Excel.
                    MethodInfo method = methods[dispIdMember];
                    object result = method.Invoke(customRibbon, BindArguments(method, pDispParams));

                    if (method.ReturnType != typeof(void))
                        Dispatcher.SetResult(pVarResult, ConvertReturnValue(result));
                }
                else if (dispIdMember == methods.Length && pDispParams.cArgs == 1)
                {
                    // The built-in LoadImage callback - loads an embedded manifest resource named by the ribbon control's image attribute.
                    string resourceName = pDispParams.rgvarg[0].Value as string;
                    Dispatcher.SetResult(pVarResult, new DispatchObject(Picture.LoadAsIPictureDisp(LoadCustomRibbonResource(resourceName))));
                }
            }
            catch (Exception ex)
            {
                // A callback exception (from the add-in's callback, or an unsupported return/argument type) must not be
                // allowed to cross the COM boundary and crash Excel. Log it through the Excel-DNA diagnostic channel so
                // a misdeclared or throwing ribbon callback is diagnosable (rather than silently doing nothing).
                ExcelDna.Logging.Logger.ComAddIn.Error(ex, "Ribbon callback (dispId {0}) failed", dispIdMember);
            }

            return 0;
        }

        // Build the managed argument array for a ribbon callback from the Excel-supplied DispParams,
        // converting each variant to the type expected by the callback parameter.
        private static object[] BindArguments(MethodInfo method, in DispParams pDispParams)
        {
            ParameterInfo[] parameters = method.GetParameters();
            object[] args = new object[parameters.Length];

            int argCount = pDispParams.rgvarg?.Length ?? 0;
            for (int i = 0; i < parameters.Length; ++i)
            {
                object value = i < argCount ? pDispParams.rgvarg[i].Value : null;
                args[i] = ConvertArgument(value, parameters[i].ParameterType);
            }

            return args;
        }

        private static object ConvertArgument(object value, Type parameterType)
        {
            // The first argument of most ribbon callbacks is the IRibbonControl - we present it as a RibbonControl wrapper.
            if (parameterType == typeof(CustomUI.RibbonControl))
                return new CustomUI.RibbonControl((value as DispatchObject)?.ComObject as IRibbonControl);

            // The onLoad callback receives the IRibbonUI - we present it as a RibbonUI wrapper so the add-in can call Invalidate().
            if (parameterType == typeof(CustomUI.RibbonUI))
                return value is DispatchObject dispatchObject ? new CustomUI.RibbonUI(dispatchObject) : null;

            // Simple callback arguments (bool pressed, string text, int selectedIndex, ...) are passed straight through.
            return value;
        }

        private static object ConvertReturnValue(object result)
        {
            // A getImage callback may return raw image bytes - convert them to an IPictureDisp like the built-in LoadImage does.
            if (result is byte[] imageData)
                return new DispatchObject(Picture.LoadAsIPictureDisp(imageData));

            // bool / string / int / double / enum return values are marshalled directly by the VariantMarshaller.
            return result;
        }

        #region IDTExtensibility2 interface
        public virtual void OnConnection(IntPtr Application, ext_ConnectMode ConnectMode, IntPtr AddInInst, ref Generator.Interfaces.SafeArray custom)
        {
        }

        public virtual void OnDisconnection(ext_DisconnectMode RemoveMode, ref Generator.Interfaces.SafeArray custom)
        {
            ExcelComAddInHelper.OnUnloadComAddIn(this, null);
        }

        public virtual void OnAddInsUpdate(ref Generator.Interfaces.SafeArray custom)
        {
        }

        public virtual void OnStartupComplete(ref Generator.Interfaces.SafeArray custom)
        {
        }

        public virtual void OnBeginShutdown(ref Generator.Interfaces.SafeArray custom)
        {
        }

        public int GetCustomUI([MarshalAs(UnmanagedType.BStr)] string RibbonID, [MarshalAs(UnmanagedType.BStr)] out string result)
        {
            result = customRibbon.GetCustomUI(RibbonID);
            return 0;
        }
        #endregion

        private byte[] LoadCustomRibbonResource(string name)
        {
            using (var stream = customRibbon.GetType().Assembly.GetManifestResourceStream(name))
            {
                using (System.IO.MemoryStream memoryStream = new())
                {
                    stream.CopyTo(memoryStream);
                    return memoryStream.ToArray();
                }
            }
        }
    }
}

#endif
