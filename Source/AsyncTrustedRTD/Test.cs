using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AsyncTrustedRTD
{
    public class Test
    {
        [ExcelDna.Integration.ExcelFunction(Description = "My first .NET function")]
        public static string SayHello(string name)
        {
            return "Hello " + name;
        }

        [ExcelDna.Integration.ExcelFunction(Description = "My first .NET function")]
        public static object SayHelloAsync(string name)
        {
            return ExcelDna.Integration.ExcelAsyncUtil.Run("RunSomethingDelay", null, () => RunSomethingDelay());
        }

        public static string RunSomethingDelay()
        {
            System.Threading.Thread.Sleep(2000);
            return "test";
        }
    }
    
}
