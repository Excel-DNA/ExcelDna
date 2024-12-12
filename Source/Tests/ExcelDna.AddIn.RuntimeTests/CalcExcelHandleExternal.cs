using ExcelDna.Integration;

[assembly: ExcelHandleExternal(typeof(ExcelDna.AddIn.RuntimeTests.CalcExcelHandleExternal))]

namespace ExcelDna.AddIn.RuntimeTests
{
    public class CalcExcelHandleExternal
    {
        private double d1, d2;

        public CalcExcelHandleExternal(double d1, double d2)
        {
            this.d1 = d1;
            this.d2 = d2;
        }

        public double Mul()
        {
            return d1 * d2;
        }
    }
}
