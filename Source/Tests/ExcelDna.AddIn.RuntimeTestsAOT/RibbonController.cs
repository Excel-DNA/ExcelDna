using ExcelDna.Integration.CustomUI;

namespace ExcelDna.AddIn.RuntimeTestsAOT
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
              <button id='button1' label='My Native Button' onAction='OnButtonPressed'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            MessageBox.Show("Hello from native control " + control.Id);
        }
    }
}

