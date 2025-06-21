#if COM_GENERATED

#nullable enable

namespace ExcelDna.Integration.ComInterop.Generator
{
    internal class DynamicComObject : IDynamic
    {
        private Interfaces.DispatchObject dispatchObject;

        public DynamicComObject(Interfaces.DispatchObject dispatchObject)
        {
            this.dispatchObject = dispatchObject;
        }

        public object? GetProperty(string name)
        {
            return WrapDispatch(dispatchObject.GetProperty(name));
        }

        public T GetProperty<T>(string name)
        {
            return (T)GetProperty(name)!;
        }

        public object? GetProperty(string name, object[]? args)
        {
            return WrapDispatch(dispatchObject.GetProperty(name, args));
        }

        public T GetProperty<T>(string name, object[]? args)
        {
            return (T)GetProperty(name, args)!;
        }

        public void SetProperty(string name, object value)
        {
            dispatchObject.SetProperty(name, value);
        }

        public object? this[int index]
        {
            get
            {
                return WrapDispatch(dispatchObject.GetIndex(index));
            }
        }

        public object? this[string index]
        {
            get
            {
                return WrapDispatch(dispatchObject.GetIndex(index));
            }
        }

        public object? Invoke(string name, object[] args)
        {
            return WrapDispatch(dispatchObject.Invoke(name, args));
        }

        private static object? WrapDispatch(object? o)
        {
            if (o is Interfaces.DispatchObject dispatch)
                return new DynamicComObject(dispatch);

            return o;
        }
    }
}

#endif
