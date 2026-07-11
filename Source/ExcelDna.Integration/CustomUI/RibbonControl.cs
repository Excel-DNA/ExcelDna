#if COM_GENERATED

using System.Runtime.InteropServices;

namespace ExcelDna.Integration.CustomUI
{
    public class RibbonControl
    {
        private ComInterop.Generator.Interfaces.IRibbonControl control;

        public string Id
        {
            get
            {
                control.get_Id(out string result);
                return result;
            }
        }

        public string Tag
        {
            get
            {
                control.get_Tag(out string result);
                return result;
            }
        }

        public object Context
        {
            get
            {
                int hr = control.get_Context(out nint result);
                Marshal.ThrowExceptionForHR(hr);
                if (result == 0)
                    return null;

                return new ComInterop.Generator.DynamicComObject(new ComInterop.Generator.Interfaces.DispatchObject(result));
            }
        }

        internal RibbonControl(ComInterop.Generator.Interfaces.IRibbonControl control)
        {
            this.control = control;
        }
    }
}

#endif
