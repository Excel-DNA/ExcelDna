using ExcelDna.Integration;

namespace ExcelDna.AddIn.RuntimeTestsAOT
{
    public class Commands
    {
        [ExcelCommand(MenuText = "MyNativeCommand")]
        public static void NativeCommand()
        {
            System.Diagnostics.Trace.WriteLine("RuntimeTestsAOT NativeCommand");
        }
    }
}
