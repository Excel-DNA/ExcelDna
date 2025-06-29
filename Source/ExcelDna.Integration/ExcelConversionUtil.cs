namespace ExcelDna.Integration
{
    public class ExcelConversionUtil
    {
#if COM_GENERATED
        public static IRange ReferenceToRange(ExcelReference reference)
        {
            IDynamic app = ExcelDnaUtil.DynamicApplication;

            string sheetName = (string)XlCall.Excel(XlCall.xlSheetNm, reference);
            int index = sheetName.LastIndexOf("]");
            sheetName = sheetName.Substring(index + 1);
            IDynamic ws = (IDynamic)app.Get<IDynamic>("Sheets")[sheetName];

            IDynamic target = app.Get<IDynamic>("Range", [
                ws.Get<IDynamic>("Cells", [reference.RowFirst + 1, reference.ColumnFirst + 1]),
                ws.Get<IDynamic>("Cells", [reference.RowLast + 1, reference.ColumnLast + 1])
                ]);

            for (int iInnerRef = 1; iInnerRef < reference.InnerReferences.Count; iInnerRef++)
            {
                ExcelReference innerRef = reference.InnerReferences[iInnerRef];
                IDynamic innerTarget = app.Get<IDynamic>("Range", [
                    ws.Get<IDynamic>("Cells", [innerRef.RowFirst + 1, innerRef.ColumnFirst + 1]),
                    ws.Get<IDynamic>("Cells", [innerRef.RowLast + 1, innerRef.ColumnLast + 1])
                    ]);
                target = app.Invoke<IDynamic>("Union", [target, innerTarget]);
            }

            return new ComInterop.Generator.Range(target as ComInterop.Generator.DynamicComObject);
        }
#else
        public static Microsoft.Office.Interop.Excel.Range ReferenceToRange(ExcelReference reference)
        {
            Microsoft.Office.Interop.Excel.Application app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;

            string sheetName = (string)XlCall.Excel(XlCall.xlSheetNm, reference);
            int index = sheetName.LastIndexOf("]");
            sheetName = sheetName.Substring(index + 1);
            Microsoft.Office.Interop.Excel.Worksheet ws = app.Sheets[sheetName];
            Microsoft.Office.Interop.Excel.Range target = app.Range[ws.Cells[reference.RowFirst + 1, reference.ColumnFirst + 1], ws.Cells[reference.RowLast + 1, reference.ColumnLast + 1]];

            for (int iInnerRef = 1; iInnerRef < reference.InnerReferences.Count; iInnerRef++)
            {
                ExcelReference innerRef = reference.InnerReferences[iInnerRef];
                Microsoft.Office.Interop.Excel.Range innerTarget = app.Range[ws.Cells[innerRef.RowFirst + 1, innerRef.ColumnFirst + 1], ws.Cells[innerRef.RowLast + 1, innerRef.ColumnLast + 1]];
                target = app.Union(target, innerTarget);
            }

            return target;
        }
#endif
    }
}
