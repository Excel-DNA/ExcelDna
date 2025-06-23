using System;

namespace ExcelDna.Integration.ComInterop
{
    internal interface IType
    {
        object GetObject(IntPtr pUnk);
        void ReleaseObject(object comObject);

        bool HasProperty(string name, object comObject);
        object GetProperty(string name, object comObject);
        void SetProperty(string name, object value, object comObject);
        object GetIndex(int i, object comObject);
        object GetIndex(string name, object comObject);
        object Invoke(string name, object[] args, object comObject);
        bool Is(Guid guid, object comObject);
        IntPtr QueryInterface(Guid guid, object comObject);
    }
}
