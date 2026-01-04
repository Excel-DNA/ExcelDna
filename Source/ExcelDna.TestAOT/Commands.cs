using ExcelDna.Integration;

namespace ExcelDna.TestAOT
{
    public class Commands
    {
        [ExcelCommand(MenuText = "NativeTestCommand")]
        public static void NativeTestCommand()
        {
            System.Diagnostics.Trace.WriteLine("ExcelDna.Test NativeTestCommand");
        }

        [ExcelCommand(MenuText = "NativeQueueMacro")]
        public static void NativeQueueMacro()
        {
            ExcelAsyncUtil.QueueMacro("NativeTestCommand");
        }
    }
}
