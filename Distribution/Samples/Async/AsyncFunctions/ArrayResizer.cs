using System.Collections.Generic;
using ExcelDna.Integration;
using System.Diagnostics;

namespace AsyncFunctions
{
    public class ArrayResizer
    {
        static Queue<ExcelReference> ResizeJobs = new Queue<ExcelReference>();

        // This function will run in the UDF context.
        public static object Resize(object[,] array)
        {
            ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            return Resize(array, caller);
        }

        internal static object ResizeObservable(object[,] array, ExcelReference caller)
        {
            object callerAfter = XlCall.Excel(XlCall.xlfCaller);
            if (callerAfter == null)
            {
                // This is the good RTD array call
                return Resize(array, caller);
            }
            // Some spurious RTD call - just return.
            return array;
        }

        public static object Resize(object[,] array, ExcelReference caller)
        {
            if (caller == null)
            {
                Debug.Print("Resize - Abandoning - No Caller");
                return array;
            }

            int rows = array.GetLength(0);
            int columns = array.GetLength(1);

            if ((caller.RowLast - caller.RowFirst + 1 != rows) ||
                (caller.ColumnLast - caller.ColumnFirst + 1 != columns))
            {
                // Size problem: enqueue job, call async update and return #N/A
                EnqueueResize(caller, rows, columns);
                ExcelAsyncUtil.QueueAsMacro(DoResizing);
            }

            // Size is already OK - just return result
            return array;
        }

        static void EnqueueResize(ExcelReference caller, int rows, int columns)
        {
            ExcelReference target = new ExcelReference(caller.RowFirst, caller.RowFirst + rows - 1, caller.ColumnFirst, caller.ColumnFirst + columns - 1, caller.SheetId);
            ResizeJobs.Enqueue(target);
        }

        static void DoResizing()
        {
            while (ResizeJobs.Count > 0)
            {
                DoResize(ResizeJobs.Dequeue());
            }
        }

        static void DoResize(ExcelReference target)
        {
            object oldEcho = XlCall.Excel(XlCall.xlfGetWorkspace, 40);
            object oldCalculationMode = XlCall.Excel(XlCall.xlfGetDocument, 14);
            try
            {
                // Get the current state for reset later
                XlCall.Excel(XlCall.xlcEcho, false);
                XlCall.Excel(XlCall.xlcOptionsCalculation, 3);

                // Get the formula in the first cell of the target
                string formula = (string)XlCall.Excel(XlCall.xlfGetCell, 41, target);
                ExcelReference firstCell = new ExcelReference(target.RowFirst, target.RowFirst, target.ColumnFirst, target.ColumnFirst, target.SheetId);

                bool isFormulaArray = (bool)XlCall.Excel(XlCall.xlfGetCell, 49, target);
                if (isFormulaArray)
                {
                    object oldSelectionOnActiveSheet = XlCall.Excel(XlCall.xlfSelection);
                    object oldActiveCell = XlCall.Excel(XlCall.xlfActiveCell);

                    // Remember old selection and select the first cell of the target
                    string firstCellSheet = (string)XlCall.Excel(XlCall.xlSheetNm, firstCell);
                    XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { firstCellSheet });
                    object oldSelectionOnArraySheet = XlCall.Excel(XlCall.xlfSelection);
                    XlCall.Excel(XlCall.xlcFormulaGoto, firstCell);

                    // Extend the selection to the whole array and clear
                    XlCall.Excel(XlCall.xlcSelectSpecial, 6);
                    ExcelReference oldArray = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);

                    oldArray.SetValue(ExcelEmpty.Value);
                    XlCall.Excel(XlCall.xlcSelect, oldSelectionOnArraySheet);
                    XlCall.Excel(XlCall.xlcFormulaGoto, oldSelectionOnActiveSheet);
                }
                // Get the formula and convert to R1C1 mode
                bool isR1C1Mode = (bool)XlCall.Excel(XlCall.xlfGetWorkspace, 4);
                string formulaR1C1 = formula;
                if (!isR1C1Mode)
                {
                    // Set the formula into the whole target
                    formulaR1C1 = (string)XlCall.Excel(XlCall.xlfFormulaConvert, formula, true, false, ExcelMissing.Value, firstCell);
                }
                // Must be R1C1-style references
                object ignoredResult;
                //Debug.Print("Resizing START: " + target.RowLast);
                XlCall.XlReturn retval = XlCall.TryExcel(XlCall.xlcFormulaArray, out ignoredResult, formulaR1C1, target);
                //Debug.Print("Resizing FINISH");

                // TODO: Dummy action to clear the undo stack
                
                if (retval != XlCall.XlReturn.XlReturnSuccess)
                {
                    // TODO: Consider what to do now!?
                    // Might have failed due to array in the way.
                    firstCell.SetValue("'" + formula);
                }
            }
            finally
            {
                XlCall.Excel(XlCall.xlcEcho, oldEcho);
                XlCall.Excel(XlCall.xlcOptionsCalculation, oldCalculationMode);
            }
        }
    }
}