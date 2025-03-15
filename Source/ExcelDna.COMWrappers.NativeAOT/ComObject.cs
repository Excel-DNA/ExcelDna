using Addin.ComApi;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using System.Runtime.InteropServices.ComTypes;

namespace ExcelDna.COMWrappers.NativeAOT
{
    internal class ComObject
    {
        private IntPtr pUnk;
        private IDispatch? dispatch;
        private Guid emptyGuid = Guid.Empty;

        private const int LOCALE_USER_DEFAULT = 0x0400;
        private const int DISPID_PROPERTYPUT = -3;


        public ComObject(IntPtr pUnk)
        {
            this.pUnk = pUnk;

            ComWrappers cw = new StrategyBasedComWrappers();
            dispatch = cw.GetOrCreateObjectForComInstance(pUnk, CreateObjectFlags.None) as IDispatch;
        }

        public bool HasProperty(string name)
        {
            try
            {
                GetDispIDs(name);
                return true;
            }
            catch
            {
            }

            return false;
        }

        public object? GetProperty(string name)
        {
            Addin.Types.Managed.DispParams dispParams = new();

            return InvokeWrapper(name, INVOKEKIND.INVOKE_PROPERTYGET, dispParams);
        }

        public void SetProperty(string name, object value)
        {
            var dispParams = new Addin.Types.Managed.DispParams
            {
                rgvarg = [new Addin.Types.Managed.Variant(value)],
                rgdispidNamedArgs = DISPID_PROPERTYPUT,
                cArgs = 1,
                cNamedArgs = 1
            };

            InvokeWrapper(name, INVOKEKIND.INVOKE_PROPERTYPUT, dispParams);
        }

        public object? GetIndex(int i)
        {
            var index = i;

            var dispParams = new Addin.Types.Managed.DispParams
            {
                rgvarg = [new Addin.Types.Managed.Variant(index)],
                rgdispidNamedArgs = 0,
                cArgs = 1,
                cNamedArgs = 0
            };

            return InvokeWrapper("Item", INVOKEKIND.INVOKE_PROPERTYGET, dispParams);
        }

        public object? Invoke(string name, object[] args)
        {
            Addin.Types.Managed.Variant[] a = new Addin.Types.Managed.Variant[args.Length];
            for (int i = 0; i < args.Length; ++i)
            {
                object? o = args[i].GetType().IsEnum ? (int)args[i] : args[i];
                a[i] = new Addin.Types.Managed.Variant(o);
            }

            var dispParams = new Addin.Types.Managed.DispParams
            {
                rgvarg = a,
                rgdispidNamedArgs = 0,
                cArgs = a.Length,
                cNamedArgs = 0
            };

            return InvokeWrapper(name, INVOKEKIND.INVOKE_FUNC, dispParams);
        }

        public unsafe bool HasInterface(ref Guid guid)
        {
            IIUnknownStrategy iUnknownStrategy = StrategyBasedComWrappers.DefaultIUnknownStrategy;
            iUnknownStrategy.QueryInterface(pUnk.ToPointer(), in guid, out void* ppObj);
            return ppObj != null;
        }

        private int[] GetDispIDs(string propName)
        {
            var names = new string[] { propName };

            var dispIds = new int[names.Length];
            var hr = dispatch!.GetIDsOfNames(
                ref emptyGuid,
                names,
                (uint)names.Length,
                LOCALE_USER_DEFAULT,
                dispIds
            );

            if (hr < 0)
            {
                Marshal.ThrowExceptionForHR(hr);
            }

            return dispIds;
        }

        private object? InvokeWrapper(string propName, INVOKEKIND kind, Addin.Types.Managed.DispParams dispParams)
        {
            var dispIds = GetDispIDs(propName);

            Addin.Types.Managed.ExcepInfo pExcepInfo = new();
            Addin.Types.Managed.Variant pVarResult = new();
            uint puArgErr = 0;

            var hr = dispatch!.Invoke(
                dispIds[0],
                emptyGuid,
                LOCALE_USER_DEFAULT,
                kind,
                ref dispParams,
                ref pVarResult,
                ref pExcepInfo,
                ref puArgErr
            );

            Marshal.ThrowExceptionForHR(hr);

            if (pVarResult.Value is IDispatch interfacePtr)
            {
                return new ComObject(pVarResult.DispVal);
            }

            return pVarResult.Value;
        }
    }
}
