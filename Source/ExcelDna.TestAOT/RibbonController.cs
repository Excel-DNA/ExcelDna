using ExcelDna.Integration.CustomUI;

namespace ExcelDna.TestAOT
{
    // Dynamic ribbon test fixture for the NativeAOT path.
    //
    // Exercises the ribbon callback behaviours reported as broken under NativeAOT
    // (https://groups.google.com/g/exceldna/c/qwQLiufM5d4):
    //   - onLoad receives a usable IRibbonUI (here the RibbonUI wrapper) so that Invalidate() works
    //   - getLabel / getEnabled / getVisible callbacks return a value back to Excel
    //   - a toggleButton onAction with the (control, pressed) 2-argument signature
    //   - loadImage='LoadImage' driven image loading from embedded resources
    public class RibbonController : IExcelRibbon
    {
        private RibbonUI? ribbon;
        private bool toggledOn = false;

        public string GetCustomUI(string RibbonID)
        {
            return @"
        <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' loadImage='LoadImage' onLoad='OnLoad'>
            <ribbon>
                <tabs>
                    <tab id='tab1' label='My Native Tab'>
                    <group id='group1' label='My Native Group'>
                        <button id='button1' tag='tagdata1' label='My Native Button1' image='ExcelDna.TestAOT.app.ico' onAction='OnButtonPressed1'/>
                        <button id='button2' tag='tagdata2' label='My Native Button2' image='ExcelDna.TestAOT.Package.ico' onAction='OnButtonPressed2'/>
                        <toggleButton id='toggle1' getLabel='GetToggleLabel' getPressed='GetTogglePressed' onAction='OnToggle'/>
                        <button id='dynamicButton' getLabel='GetDynamicLabel' getEnabled='GetDynamicEnabled' getVisible='GetDynamicVisible' onAction='OnInvalidate'/>
                    </group >
                    </tab>
                </tabs>
            </ribbon>
        </customUI>";
        }

        // onLoad - Excel passes the IRibbonUI here. Keep a reference so we can Invalidate() later.
        public void OnLoad(RibbonUI ribbon)
        {
            this.ribbon = ribbon;
            System.Diagnostics.Trace.WriteLine("OnLoad - native ribbon loaded");
        }

        public void OnButtonPressed1(RibbonControl control)
        {
            System.Diagnostics.Trace.WriteLine($"Hello1 from native control id={control.Id} tag={control.Tag}");
        }

        public void OnButtonPressed2(RibbonControl control)
        {
            System.Diagnostics.Trace.WriteLine($"Hello2 from native control id={control.Id} tag={control.Tag}");
        }

        // toggleButton callbacks - getLabel / getPressed return values, onAction has the 2-argument (control, pressed) shape.
        public string GetToggleLabel(RibbonControl control) => toggledOn ? "Toggled ON" : "Toggled OFF";

        public bool GetTogglePressed(RibbonControl control) => toggledOn;

        public void OnToggle(RibbonControl control, bool pressed)
        {
            toggledOn = pressed;
            System.Diagnostics.Trace.WriteLine($"OnToggle id={control.Id} pressed={pressed}");
            ribbon?.InvalidateControl("dynamicButton");
        }

        // Dynamic button callbacks - their values change based on toggle state and are refreshed via Invalidate().
        public string GetDynamicLabel(RibbonControl control) => toggledOn ? "Enabled - click to refresh" : "Disabled while toggle is off";

        public bool GetDynamicEnabled(RibbonControl control) => toggledOn;

        public bool GetDynamicVisible(RibbonControl control) => true;

        public void OnInvalidate(RibbonControl control)
        {
            System.Diagnostics.Trace.WriteLine("OnInvalidate - invalidating whole ribbon");
            ribbon?.Invalidate();
        }
    }
}
