using System.Windows.Forms;
using ExcelDna.Integration;

public class TestAddIn
{
    [ExcelCommand(MenuName="Clear", MenuText="Clear A2:B5")]
    public static void ClearA2B5()
    {
        ExcelReference xlRef = new ExcelReference(1, 4, 0, 1);
        int rows = xlRef.RowLast - xlRef.RowFirst + 1;
        int cols = xlRef.ColumnLast - xlRef.ColumnFirst + 1;
        object[,] values = new object[rows, cols]; // nulls
        xlRef.SetValue(values);

        MessageBox.Show("Done clearing!");
    }

    public static string HelloFromPackedSource()
    {
        return "Hello from Packed Source!";
    }
}