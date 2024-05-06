namespace ExcelDna.Integration
{
    public class ExcelObjectHandle
    {
        public object Object { get; }

        public ExcelObjectHandle(object o)
        {
            Object = o;
        }
    }
}
