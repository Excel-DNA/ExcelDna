#if COM_GENERATED

using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using System.Runtime.InteropServices.ComTypes;
using System;
using System.Linq;

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
{
    internal class DispatchObject : UnknownObject
    {
        private IDispatch dispatch;
        private Guid emptyGuid = Guid.Empty;

        private const int LOCALE_USER_DEFAULT = 0x0400;
        private const int DISPID_PROPERTYPUT = -3;

        public DispatchObject(IntPtr unknown) : base(unknown)
        {
            ComWrappers cw = new StrategyBasedComWrappers();
            IDispatch? dispatch = cw.GetOrCreateObjectForComInstance(P, CreateObjectFlags.None) as IDispatch;
            if (dispatch == null)
                throw new ArgumentException();

            this.dispatch = dispatch;
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
            DispParams dispParams = new();

            return InvokeWrapper(name, INVOKEKIND.INVOKE_PROPERTYGET, dispParams);
        }

        public void SetProperty(string name, object value)
        {
            var dispParams = new DispParams
            {
                rgvarg = [new Variant(value)],
                rgdispidNamedArgs = DISPID_PROPERTYPUT,
                cArgs = 1,
                cNamedArgs = 1
            };

            InvokeWrapper(name, INVOKEKIND.INVOKE_PROPERTYPUT, dispParams);
        }

        public object? GetIndex(int i)
        {
            var dispParams = new DispParams
            {
                rgvarg = [new Variant(i)],
                rgdispidNamedArgs = 0,
                cArgs = 1,
                cNamedArgs = 0
            };

            return InvokeWrapper("Item", INVOKEKIND.INVOKE_PROPERTYGET, dispParams);
        }

        public object? GetIndex(string name)
        {
            var dispParams = new DispParams
            {
                rgvarg = [new Variant(name)],
                rgdispidNamedArgs = 0,
                cArgs = 1,
                cNamedArgs = 0
            };

            return InvokeWrapper("Item", INVOKEKIND.INVOKE_PROPERTYGET, dispParams);
        }

        public object? Invoke(string name, object[] args)
        {
            DispParams dispParams;
            if (args != null)
            {
                Variant[] a = args.Select(i => new Variant(i)).ToArray();

                dispParams = new DispParams
                {
                    rgvarg = a,
                    rgdispidNamedArgs = 0,
                    cArgs = a.Length,
                    cNamedArgs = 0
                };
            }
            else
            {
                dispParams = new DispParams();
            }

            return InvokeWrapper(name, INVOKEKIND.INVOKE_FUNC, dispParams);
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

            Marshal.ThrowExceptionForHR(hr);

            return dispIds;
        }

        private object? InvokeWrapper(string propName, INVOKEKIND kind, DispParams dispParams)
        {
            var dispIds = GetDispIDs(propName);

            ExcepInfo pExcepInfo = new();
            Variant pVarResult = new();
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

            return pVarResult.Value;
        }
    }
}

#endif
