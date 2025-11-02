using ExcelDna.Integration.CustomUI;

namespace AOTRibbon
{
    public class RibbonController : IExcelRibbon
    {
        public string GetCustomUI(string RibbonID)
        {
            return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
      <ribbon>
        <tabs>
          <tab id='tab1' label='My Native Tab'>
            <group id='group1' label='My Native Group'>
              <button id='button1' tag='tagdata1' label='My Native Button1' onAction='OnButtonPressed1'/>
              <button id='button2' tag='tagdata2' label='My Native Button2' onAction='OnButtonPressed2'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        public void OnButtonPressed1(RibbonControl control)
        {
            System.Diagnostics.Trace.WriteLine($"Hello1 from native control id={control.Id} tag={control.Tag}");
        }

        public void OnButtonPressed2(RibbonControl control)
        {
            System.Diagnostics.Trace.WriteLine($"Hello2 from native control id={control.Id} tag={control.Tag}");
        }
    }
}

