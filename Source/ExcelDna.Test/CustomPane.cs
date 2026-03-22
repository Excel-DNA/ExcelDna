using ExcelDna.Integration.CustomUI;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelDna.Test
{
    internal class CustomPane
    {
        public static void Show()
        {
            var myControl = new MyUserControl();
            var customPane = CustomTaskPaneFactory.CreateCustomTaskPane(myControl, nameof(myControl),
                new Guid("dfdd066f-a8ce-4be0-ac13-20a185333473"), "1a7ad958-f8f5-43d4-9161-5bbab6ecda62");

            customPane.Visible = true;
        }
    }

    public interface IMyUserControl { }

    [ComVisible(true)]
    [Guid("c5a18d1b-b798-49cf-9a3f-37a094905170")]
    [ComDefaultInterface(typeof(IMyUserControl))]
    public class MyUserControl : UserControl, IMyUserControl { }
}
