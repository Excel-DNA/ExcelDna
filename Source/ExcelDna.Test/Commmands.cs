using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDna.Test
{
    public class Commands
    {
        [ExcelCommand(MenuText = "MyTestCommand")]
        public static void MyTestCommand()
        {
            System.Diagnostics.Trace.WriteLine("ExcelDna.Test MyTestCommand");
        }

        [ExcelCommand(MenuText = "MyQueueMacro")]
        public static void MyQueueMacro()
        {
            ExcelAsyncUtil.QueueMacro("MyTestCommand");
        }
    }
}
