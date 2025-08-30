using ExcelDna.Integration;

namespace ExcelDna.AddIn.RuntimeTestsAOT
{
    [ExcelHandle]
    public class CalcExcelHandle
    {
        private double d1, d2;

        public CalcExcelHandle(double d1, double d2)
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
