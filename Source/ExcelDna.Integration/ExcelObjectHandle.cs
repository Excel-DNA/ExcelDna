namespace ExcelDna.Integration
{
    public class ExcelObjectHandle<T>
    {
        public T Object { get; }

        public ExcelObjectHandle(T o)
        {
            Object = o;
        }
    }
}
