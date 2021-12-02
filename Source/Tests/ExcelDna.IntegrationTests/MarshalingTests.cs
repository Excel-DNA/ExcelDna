using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelDna.IntegrationTests
{
    public class MarshalingTests
    {
        [ExcelFunction]
        public static double testAdd(double d1, double d2)
        {
            return d1+d2;
        }

        [ExcelFunction]
        public static double testMult(double d1, double d2)
        {
            return d1*d2;
        }

        [ExcelFunction]
        public static object testReturnDoubleArray()
        {
            return new double[,] { { 1, 2 }, { 3, 4 } };
        }
    }
}
