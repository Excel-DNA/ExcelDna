using ExcelDna.ComInterop;

namespace ExcelDna.COMWrappers.NativeAOT
{
    public class TypeAdapter : IType
    {
        public object GetProperty(string name, object comObject)
        {
            return (comObject as ComObject)!.GetProperty(name)!;
        }

        public object GetIndex(int i, object comObject)
        {
            return (comObject as ComObject)!.GetIndex(i)!;
        }

        public bool Is(ref Guid guid, object comObject)
        {
            return (comObject as ComObject)!.HasInterface(ref guid);
        }

        public object Invoke(string name, object[] args, object comObject)
        {
            return (comObject as ComObject)!.Invoke(name, args)!;
        }

        public void SetProperty(string name, object value, object comObject)
        {
            (comObject as ComObject)!.SetProperty(name, value);
        }

        public object GetObject(IntPtr pUnk)
        {
            return new ComObject(pUnk);
        }

        public void ReleaseObject(object comObject)
        {
        }

        public bool HasProperty(string name, object comObject)
        {
            return (comObject as ComObject)!.HasProperty(name);
        }
    }
}
