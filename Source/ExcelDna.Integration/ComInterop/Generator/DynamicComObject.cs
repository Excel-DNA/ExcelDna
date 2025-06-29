#if COM_GENERATED

#nullable enable

using System.Linq;

namespace ExcelDna.Integration.ComInterop.Generator
{
    internal class DynamicComObject : IDynamic
    {
        private Interfaces.DispatchObject dispatchObject;

        public DynamicComObject(Interfaces.DispatchObject dispatchObject)
        {
            this.dispatchObject = dispatchObject;
        }

        public DynamicComObject(DynamicComObject o) : this(o.dispatchObject)
        {
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
            return WrapDispatch(dispatchObject.GetProperty(propertyName, UnwrapDispatch(args)));
        }

        public T Get<T>(string propertyName, object[]? args)
        {
            return (T)Get(propertyName, UnwrapDispatch(args))!;
        }

        public void Set(string propertyName, object value)
        {
            dispatchObject.SetProperty(propertyName, UnwrapDispatch(value));
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
            return WrapDispatch(dispatchObject.Invoke(functionName, UnwrapDispatch(args)));
        }

        public T Invoke<T>(string functionName, object[]? args)
        {
            return (T)Invoke(functionName, UnwrapDispatch(args))!;
        }

        private static object? WrapDispatch(object? o)
        {
            if (o is Interfaces.DispatchObject dispatch)
                return new DynamicComObject(dispatch);

            return o;
        }

        private static object UnwrapDispatch(object o)
        {
            if (o is DynamicComObject dynamicComObject)
                return dynamicComObject.dispatchObject;

            return o;
        }

        private static object[]? UnwrapDispatch(object[]? args)
        {
            if (args == null)
                return null;

            return args.Select(i => UnwrapDispatch(i)).ToArray();
        }
    }
}

#endif
