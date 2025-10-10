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
            if (dispIdMember >= 0 && dispIdMember < methods.Length && pDispParams.cArgs == 1)
            {
                CustomUI.RibbonControl ribbonControl = new((pDispParams.rgvarg[0].Value as DispatchObject).ComObject as IRibbonControl);
                methods[dispIdMember].Invoke(customRibbon, [ribbonControl]);
            }

            if (dispIdMember == methods.Length && pDispParams.cArgs == 1)
            {
                string resourceName = pDispParams.rgvarg[0].Value as string;
                Dispatcher.SetResult(pVarResult, new DispatchObject(Picture.LoadAsIPictureDisp(LoadCustomRibbonResource(resourceName))));
            }

            return 0;
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
