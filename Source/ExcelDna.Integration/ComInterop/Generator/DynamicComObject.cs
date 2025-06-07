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
            return dispatchObject.GetProperty(name);
        }
    }
}

#endif
