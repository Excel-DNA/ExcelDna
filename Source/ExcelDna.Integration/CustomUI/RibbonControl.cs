#if COM_GENERATED

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

        internal RibbonControl(ComInterop.Generator.Interfaces.IRibbonControl control)
        {
            this.control = control;
        }
    }
}

#endif
