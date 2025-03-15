using Addin.Types.Managed;
using System.Dynamic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using static Addin.Types.Unmanaged.ComConstants;
using Unmanaged = Addin.Types.Unmanaged;

namespace Addin.ComApi;

public class ExcelObject : DynamicObject
{
    public IDispatch _interfacePtr;
    Guid emptyGuid = Guid.Empty;
    bool _verbose;

    public ExcelObject(IDispatch? interfacePtr = null)
    {
        if (interfacePtr != null)
        {
            _interfacePtr = interfacePtr;
            return;
        }

        //// The CLSID for Excel.Application (COMView.exe->CLSID table)
        //var clsid = new Guid("{00024500-0000-0000-C000-000000000046}");

        //// COMView.exe -> CLSID table -> Type column
        //var server = Unmanaged.ClsCtx.CLSCTX_LOCAL_SERVER;

        //_interfacePtr = ComClass.Create(clsid, server);
    }

    public object? GetProperty(string name)
    {
        DispParams dispParams = new();

        return InvokeWrapper(name, INVOKEKIND.INVOKE_PROPERTYGET, dispParams);
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

    public override bool TryGetMember(GetMemberBinder binder, out object? result)
    {
        DispParams dispParams = new();

        result = InvokeWrapper(binder.Name, INVOKEKIND.INVOKE_PROPERTYGET, dispParams);

        return true;
    }

    public override bool TrySetMember(SetMemberBinder binder, object value)
    {
        var dispParams = new DispParams
        {
            rgvarg = [new Variant(value)],
            rgdispidNamedArgs = DISPID_PROPERTYPUT,
            cArgs = 1,
            cNamedArgs = 1
        };

        InvokeWrapper(binder.Name, INVOKEKIND.INVOKE_PROPERTYPUT, dispParams);

        return true;
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

    public override bool TryGetIndex(GetIndexBinder binder, object[] indexes, out object? result)
    {
        var index = (int)indexes[0];

        var dispParams = new DispParams
        {
            rgvarg = [new Variant(index)],
            rgdispidNamedArgs = 0,
            cArgs = 1,
            cNamedArgs = 0
        };

        result = InvokeWrapper("Item", INVOKEKIND.INVOKE_PROPERTYGET, dispParams);

        return true;
    }

    public override bool TrySetIndex(SetIndexBinder binder, object[] indexes, object value)
    {
        throw new NotImplementedException();
    }

    public override bool TryInvokeMember(
        InvokeMemberBinder binder,
        object[] args,
        out object result
    )
    {
        //DispParams dispParams = new();

        Variant[] a = new Variant[args.Length];
        for (int i = 0; i < args.Length; ++i)
            a[i] = new Variant(args[i]);

        var dispParams = new DispParams
        {
            rgvarg = a,
            rgdispidNamedArgs = 0,
            cArgs = a.Length,
            cNamedArgs = 0
        };

        result = InvokeWrapper(binder.Name, INVOKEKIND.INVOKE_FUNC, dispParams);

        return true;
    }

    public object InvokeMember(
    string name,
    object[] args
)
    {
        //DispParams dispParams = new();

        Variant[] a = new Variant[args.Length];
        for (int i = 0; i < args.Length; ++i)
        {
            object? o = args[i].GetType().IsEnum ? (int)args[i] : args[i];

            System.Diagnostics.Trace.WriteLine("[InvokeMember] " + name + " " + o?.GetType().ToString());
            a[i] = new Variant(o);
        }

        var dispParams = new DispParams
        {
            rgvarg = a,
            rgdispidNamedArgs = 0,
            cArgs = a.Length,
            cNamedArgs = 0
        };

        return InvokeWrapper(name, INVOKEKIND.INVOKE_FUNC, dispParams);
    }

    public object InvokeWrapper(string propName, INVOKEKIND kind, DispParams dispParams)
    {
        var dispIds = GetDispIDs(propName);

        ExcepInfo pExcepInfo = new();
        Variant pVarResult = new();
        uint puArgErr = 0;

        var hr = _interfacePtr.Invoke(
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

        // Found an IDispatch object - swap current instance
        if (pVarResult.Value is IDispatch interfacePtr)
        {
            return new ExcelObject(interfacePtr);
        }

        return pVarResult.Value;
    }

    private int[] GetDispIDs(string propName)
    {
        var names = new string[] { propName };

        var dispIds = new int[names.Length];
        var hr = _interfacePtr.GetIDsOfNames(
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

        if (_verbose)
        {
            for (int i = 0; i < names.Length; i++)
                Console.WriteLine($"{names[i]}: {dispIds[i]}");
        }

        return dispIds;
    }
}
