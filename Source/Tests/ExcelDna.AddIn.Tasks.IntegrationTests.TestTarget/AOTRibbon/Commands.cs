using ExcelDna.Integration;

namespace AOTRibbon
{
    public class Commands
    {
        [ExcelCommand(MenuText = "MyNativeCommand")]
        public static void NativeCommand()
        {
            ExcelDnaUtil.DynamicApplication.Set("Caption", "My NativeCommand");
        }
    }
}
