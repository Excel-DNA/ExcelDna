using ExcelDna.Integration;

namespace AsyncFunctions
{
    public static class ArrayResizerTest
    {
        public static object MakeArray(int rows, int columns)
        {
            object[,] result = new string[rows, columns];
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    result[i, j] = string.Format("({0},{1})", i, j);
                }
            }

            return result;
        }

        public static object MakeArrayAndResize(int rows, int columns)
        {
            object result = MakeArray(rows, columns);
            // Call Resize via Excel - so if the Resize add-in is not part of this code, it should still work.
            return XlCall.Excel(XlCall.xlUDF, "Resize", result);
        }
    }
}
