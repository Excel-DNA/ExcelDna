using ExcelDna.Integration;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
    public class UI
    {
        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void Command()
        {
            CommandBarPopup? menu = FindPopupMenu("ExcelDna.AddIn.RuntimeTests add-in");
            Assert.NotNull(menu);

            CommandBarButton? button = FindButton(menu, "MyCommandHello");
            Assert.NotNull(button);

            button.Execute();

            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
            functionRange.Formula = "=MyFunctionExecutionLog()";
            Assert.True(functionRange.Value.ToString().Contains("Hello command."));
        }

        private static CommandBarPopup? FindPopupMenu(string name)
        {
            var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            var controls = app.CommandBars[1].Controls;
            for (int i = 1; i <= controls.Count; i++)
            {
                CommandBarControl control = controls[i];
                if (control.Caption == name && control is CommandBarPopup menu)
                    return menu;
            }

            return null;
        }

        private static CommandBarButton? FindButton(CommandBarPopup menu, string name)
        {
            CommandBarControls menuControls = menu.Controls;
            for (int i = 1; i <= menuControls.Count; i++)
            {
                CommandBarControl control = menuControls[i];
                if (control.Caption == name && control is CommandBarButton button)
                    return button;
            }

            return null;
        }
    }
}
