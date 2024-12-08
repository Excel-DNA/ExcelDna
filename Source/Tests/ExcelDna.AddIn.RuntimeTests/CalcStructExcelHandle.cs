using ExcelDna.Integration;

namespace ExcelDna.AddIn.RuntimeTests
{
    [ExcelHandle]
    public struct CalcStructExcelHandle
    {
        private double d1, d2;

        public CalcStructExcelHandle(double d1, double d2)
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
