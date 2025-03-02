using ExcelDna.Integration;

namespace ExcelDna.AddIn.RuntimeTestsAOT
{
    public class Commands
    {
        [ExcelCommand(MenuText = "")]
        public static void NativeCommand()
        {
            System.Diagnostics.Trace.WriteLine("RuntimeTestsAOT NativeCommand");
        }
    }
}
