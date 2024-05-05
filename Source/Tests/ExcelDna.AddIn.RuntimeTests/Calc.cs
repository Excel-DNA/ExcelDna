namespace ExcelDna.AddIn.RuntimeTests
{
    internal class Calc
    {
        private double d1, d2;

        public Calc(double d1, double d2)
        {
            this.d1 = d1;
            this.d2 = d2;
        }

        public double Sum()
        {
            return d1 + d2;
        }
    }
}
