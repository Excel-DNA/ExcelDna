using ExcelDna.Integration.Extensibility;
using ExcelDna.Logging;
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

#if COM_GENERATED
using System.Runtime.InteropServices.Marshalling;
#endif

namespace ExcelDna.Integration.ComInterop
{
#if COM_GENERATED
    [GeneratedComClass]
    internal partial class DummyComAddIn : Generator.Interfaces.IDTExtensibility2
    {
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
            throw new NotImplementedException();
        }

        public int Invoke(int dispIdMember, Guid riid, uint lcid, INVOKEKIND wFlags, [MarshalUsing(typeof(Generator.Interfaces.DispParamsMarshaller))] ref Generator.Interfaces.DispParams pDispParams, [MarshalUsing(typeof(Generator.Interfaces.VariantMarshaller))] out Generator.Interfaces.Variant pVarResult, [MarshalUsing(typeof(Generator.Interfaces.ExcepInfoMarshaller))] out Generator.Interfaces.ExcepInfo pExcepInfo, ref uint puArgErr)
        {
            throw new NotImplementedException();
        }

        #region IDTExtensibility2 interface
        public virtual void OnConnection(IntPtr Application, ext_ConnectMode ConnectMode, IntPtr AddInInst, ref Generator.Interfaces.SafeArray custom)
        {
            Logger.ComAddIn.Verbose("DummyComAddIn.OnConnection");
        }

        public virtual void OnDisconnection(ext_DisconnectMode RemoveMode, ref Generator.Interfaces.SafeArray custom)
        {
            Logger.ComAddIn.Verbose("DummyComAddIn.OnDisconnection");
        }

        public virtual void OnAddInsUpdate(ref Generator.Interfaces.SafeArray custom)
        {
            Logger.ComAddIn.Verbose("DummyComAddIn.OnAddInsUpdate");
        }

        public virtual void OnStartupComplete(ref Generator.Interfaces.SafeArray custom)
        {
            Logger.ComAddIn.Verbose("DummyComAddIn.OnStartupComplete");
        }

        public virtual void OnBeginShutdown(ref Generator.Interfaces.SafeArray custom)
        {
            Logger.ComAddIn.Verbose("DummyComAddIn.OnBeginShutdown");
        }
        #endregion
    }
#else
    [ComVisible(true)]
#pragma warning disable CS0618 // Type or member is obsolete (but probably not forever)
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
#pragma warning restore CS0618 // Type or member is obsolete
    internal class DummyComAddIn : IDTExtensibility2
    {
        #region IDTExtensibility2 interface
        public virtual void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            Logger.ComAddIn.Verbose("DummyComAddIn.OnConnection");
        }

        public virtual void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Logger.ComAddIn.Verbose("DummyComAddIn.OnDisconnection");
        }

        public virtual void OnAddInsUpdate(ref Array custom)
        {
            Logger.ComAddIn.Verbose("DummyComAddIn.OnAddInsUpdate");
        }

        public virtual void OnStartupComplete(ref Array custom)
        {
            Logger.ComAddIn.Verbose("DummyComAddIn.OnStartupComplete");
        }

        public virtual void OnBeginShutdown(ref Array custom)
        {
            Logger.ComAddIn.Verbose("DummyComAddIn.OnBeginShutdown");
        }
        #endregion
    }
#endif
}
