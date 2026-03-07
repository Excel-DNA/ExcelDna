using ExcelDna.Integration;

namespace ExcelDna.AddIn.RegistrationSample
{
    public static class MyCommands
    {
        [ExcelCommand(MenuName = "My Commands", MenuText = "Command")]
        public static void MyCommand()
        {
            return;
        }

        public static void MyGenericCommand<T>()
        {
            return;
        }
    }
}
