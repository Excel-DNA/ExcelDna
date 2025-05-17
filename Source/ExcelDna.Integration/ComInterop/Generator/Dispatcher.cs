#if COM_GENERATED

using ExcelDna.Integration.ComInterop.Generator.Interfaces;
using System;

namespace ExcelDna.Integration.ComInterop.Generator
{
    internal class Dispatcher
    {
        public class Method
        {
            public Method(string name, Action<DispParams> func)
            {
                this.Name = name;
                this.Func = func;
            }

            public string Name { get; }
            public Action<DispParams> Func { get; }
        }

        private Method[] methods;

        public Dispatcher(Method[] methods)
        {
            this.methods = methods;
        }

        public void GetIDsOfNames(string[] rgszNames, int[] rgDispId)
        {
            for (int i = 0; i < rgszNames.Length; ++i)
                rgDispId[i] = Array.FindIndex(methods, m => m.Name == rgszNames[i]);
        }

        public void Invoke(int dispIdMember, DispParams pDispParams)
        {
            if (dispIdMember >= 0 && dispIdMember < methods.Length)
            {
                methods[dispIdMember].Func.Invoke(pDispParams);
            }
        }
    }
}

#endif
