#if COM_GENERATED

using ExcelDna.Integration.ComInterop.Generator.Interfaces;

namespace ExcelDna.Integration.CustomUI
{
    // Wrapper passed to a ribbon 'onLoad' callback under NativeAOT.
    // Excel calls onLoad with the IRibbonUI object - we wrap it here so that the add-in can
    // invalidate ribbon controls (to drive dynamic getEnabled / getVisible / getLabel etc. updates).
    // The add-in should declare its onLoad callback as:  public void OnLoad(RibbonUI ribbon)
    // and keep a reference to the RibbonUI to call Invalidate() / InvalidateControl(...) later.
    public class RibbonUI
    {
        private readonly DispatchObject ribbonUI;

        internal RibbonUI(DispatchObject ribbonUI)
        {
            this.ribbonUI = ribbonUI;
        }

        public void Invalidate()
        {
            ribbonUI.Invoke("Invalidate", null);
        }

        public void InvalidateControl(string controlId)
        {
            ribbonUI.Invoke("InvalidateControl", new object[] { controlId });
        }

        public void InvalidateControlMso(string controlId)
        {
            ribbonUI.Invoke("InvalidateControlMso", new object[] { controlId });
        }

        public void ActivateTab(string controlId)
        {
            ribbonUI.Invoke("ActivateTab", new object[] { controlId });
        }

        public void ActivateTabMso(string controlId)
        {
            ribbonUI.Invoke("ActivateTabMso", new object[] { controlId });
        }

        public void ActivateTabQ(string controlId, string @namespace)
        {
            ribbonUI.Invoke("ActivateTabQ", new object[] { controlId, @namespace });
        }
    }
}

#endif
