namespace ExcelDna.Integration.CustomUI
{
    public interface ICommandBar
    {

    }

    public interface ICommandBars
    {
        ICommandBar this[string name] { get; }
    }

    public interface ICommandBarUtil
    {
        ICommandBars GetCommandBars();
    }
}
