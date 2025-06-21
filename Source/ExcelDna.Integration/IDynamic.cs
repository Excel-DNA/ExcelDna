namespace ExcelDna.Integration
{
    public interface IDynamic
    {
        object GetProperty(string name);
        T GetProperty<T>(string name);

        object GetProperty(string name, object[] args);
        T GetProperty<T>(string name, object[] args);

        void SetProperty(string name, object value);
        object this[int index] { get; }
        object this[string index] { get; }
        object Invoke(string name, object[] args);
    }
}
