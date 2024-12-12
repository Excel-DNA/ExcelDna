namespace ExcelDna.AddIn.Tasks.Utils
{
    internal interface IExcelDetector
    {
        bool TryFindLatestExcel(out string excelExePath);

        bool TryFindExcelBitness(string excelExePath, out Bitness bitness);
    }
}
