#if COM_GENERATED

using ExcelDna.Integration.Extensibility;
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
            if (cNames == 1)
            {
                System.Diagnostics.Trace.WriteLine($"ExcelRibbon.GetIDsOfNames {rgszNames[0]}");
                rgDispId[0] = Array.FindIndex(methods, i => i.Name == rgszNames[0]);
            }

            return 0;
        }

        public int Invoke(int dispIdMember, Guid riid, uint lcid, INVOKEKIND wFlags, [MarshalUsing(typeof(Generator.Interfaces.DispParamsMarshaller))] in Generator.Interfaces.DispParams pDispParams, nint pVarResult, nint pExcepInfo, ref uint puArgErr)
        {
            System.Diagnostics.Trace.WriteLine($"ExcelRibbon.Invoke {dispIdMember}");

            if (dispIdMember >= 0 && dispIdMember < methods.Length)
                methods[dispIdMember].Invoke(customRibbon, new object[] { null });

            return 0;
        }

        #region IDTExtensibility2 interface
        public virtual void OnConnection(IntPtr Application, ext_ConnectMode ConnectMode, IntPtr AddInInst, ref Generator.Interfaces.SafeArray custom)
        {
        }

        public virtual void OnDisconnection(ext_DisconnectMode RemoveMode, ref Generator.Interfaces.SafeArray custom)
        {
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
    }
}

#endif
