namespace ExcelDna.AddIn.Tasks.Utils
{
    public interface IExcelDetector
    {
        bool TryFindLatestExcel(out string excelExePath);

        bool TryFindExcelBitness(string excelExePath, out Bitness bitness);
    }
}