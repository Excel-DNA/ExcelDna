using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace CSharpAddIn
{
    public static class MyAddIn
    {
        public static string SayHello(string name)
        {
            return "Hello " + name;
        }

        [ExcelFunction(Description="A bit more than your usual adding function.")]
        public static double AddThemCS(double val1, double val2)
        {
            return val1 + val2 + 25;
        }
    }
}
