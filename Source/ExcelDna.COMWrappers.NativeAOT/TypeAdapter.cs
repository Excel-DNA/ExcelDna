using ExcelDna.COMWrappers.NativeAOT.ComInterfaces;
using ExcelDna.Integration.ComInterop;

namespace ExcelDna.COMWrappers.NativeAOT
{
    public class TypeAdapter : IType
    {
        public object GetProperty(string name, object comObject)
        {
            return (comObject as DispatchObject)!.GetProperty(name)!;
        }

        public object GetIndex(int i, object comObject)
        {
            return (comObject as DispatchObject)!.GetIndex(i)!;
        }

        public object GetIndex(string name, object comObject)
        {
            return (comObject as DispatchObject)!.GetIndex(name)!;
        }

        public bool Is(ref Guid guid, object comObject)
        {
            return (comObject as DispatchObject)!.HasInterface(ref guid);
        }

        public object Invoke(string name, object[] args, object comObject)
        {
            return (comObject as DispatchObject)!.Invoke(name, args)!;
        }

        public void SetProperty(string name, object value, object comObject)
        {
            (comObject as DispatchObject)!.SetProperty(name, value);
        }

        public object GetObject(IntPtr pUnk)
        {
            return new DispatchObject(pUnk);
        }

        public void ReleaseObject(object comObject)
        {
        }

        public bool HasProperty(string name, object comObject)
        {
            return (comObject as DispatchObject)!.HasProperty(name);
        }
    }
}
