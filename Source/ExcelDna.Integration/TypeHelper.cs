using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelDna.Integration
{
    public interface ITypeHelper
    {
        Type Type { get; }
        object CreateInstance();
        IEnumerable<MethodInfo> Methods { get; }
    }

    public class TypeHelper<T> : ITypeHelper
        where T : new()
    {
        public TypeHelper(IEnumerable<MethodInfo> methods)
        {
            this.Methods = methods;
        }

        public Type Type => typeof(T);

        public IEnumerable<MethodInfo> Methods { get; private set; }

        public object CreateInstance()
        {
            return Activator.CreateInstance<T>();
        }
    }

#if !COM_GENERATED
    public class TypeHelperDynamic : ITypeHelper
    {
        public TypeHelperDynamic(Type t)
        {
            Type = t;
            Methods = Type.GetMethods();
        }

        public Type Type { get; }

        public IEnumerable<MethodInfo> Methods { get; private set; }

        public object CreateInstance()
        {
            return Activator.CreateInstance(Type);
        }
    }
#endif
}
