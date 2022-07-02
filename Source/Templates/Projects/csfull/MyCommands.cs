using ExcelDna.Integration;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace CSfull;

public static class MyCommands
{
    // We make a command macro that can be run by:
    // * Pressing the quick menu under the Add-ins tab
    // * Pressing the shortcut key Ctrl + Shift + D
    // * Typing the name into the Alt+F8 Macro dialog (add-in macros won't we shown on this list, though)
    [ExcelCommand(MenuName = "CSfull", MenuText = "Dump Data", ShortCut = "^D")]
    public static void DumpData()
    {
        // We always get the root Application object with a call to ExcelDnaUtil.Application
        // If we reference both Window Forms, and the Excel interop assemblies, 
        // we need to be a bit careful about which 'Application' we mean.
        Application app = (Application)ExcelDnaUtil.Application;
        var newBook = app.Workbooks.Add();

        // While .NET arrays and collections are 0-based,
        // COM collections like Workbook.Sheets are 1-based.
        // We could also say newBook.Sheets["Sheet1"].
        Range targetRange = newBook.Sheets[1].Range["A1:C2"];

        // Set values to a range as a 2D object[,] array.
        object[,] newValues = new object[,] { { "One", 2, "Three" }, { true, System.DateTime.Now, "" } };
        targetRange.Value = newValues;

        // Apply some formatting, so that the time is displayed correctly
        Range dateCell = targetRange.Cells[2, 2];
        dateCell.NumberFormat = "hh:mm:ss";
    }
}

