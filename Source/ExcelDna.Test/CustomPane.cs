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
            var customPane = CustomTaskPaneFactory.CreateCustomTaskPane(myControl, nameof(myControl));

            customPane.Visible = true;
        }
    }

    public interface IMyUserControl { }

    [ComVisible(true)]
    [Guid("c5a18d1b-b798-49cf-9a3f-37a094905170")]
    [ComDefaultInterface(typeof(IMyUserControl))]
    public class MyUserControl : UserControl, IMyUserControl { }
}
