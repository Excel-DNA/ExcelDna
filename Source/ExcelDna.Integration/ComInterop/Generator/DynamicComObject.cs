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

        public object? Get(string propertyName)
        {
            return WrapDispatch(dispatchObject.GetProperty(propertyName));
        }

        public T Get<T>(string propertyName)
        {
            return (T)Get(propertyName)!;
        }

        public object? Get(string propertyName, object[]? args)
        {
            return WrapDispatch(dispatchObject.GetProperty(propertyName, args));
        }

        public T Get<T>(string propertyName, object[]? args)
        {
            return (T)Get(propertyName, args)!;
        }

        public void Set(string propertyName, object value)
        {
            dispatchObject.SetProperty(propertyName, value);
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

        public object? Invoke(string functionName, object[]? args)
        {
            return WrapDispatch(dispatchObject.Invoke(functionName, args));
        }

        public T Invoke<T>(string functionName, object[]? args)
        {
            return (T)Invoke(functionName, args)!;
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
