using ExcelDna.Integration;

namespace AOTRibbon
{
    public class Commands
    {
        [ExcelCommand(MenuText = "MyNativeCommand")]
        public static void NativeCommand()
        {
            MessageBox.Show("My NativeCommand");
        }
    }
}
