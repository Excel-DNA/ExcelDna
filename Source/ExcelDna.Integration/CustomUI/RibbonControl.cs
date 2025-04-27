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
                control.get_Id(out string id);
                return id;
            }
        }

        internal RibbonControl(ComInterop.Generator.Interfaces.IRibbonControl control)
        {
            this.control = control;
        }
    }
}

#endif
