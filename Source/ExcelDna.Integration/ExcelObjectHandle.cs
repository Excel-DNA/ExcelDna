namespace ExcelDna.Integration
{
    public class ExcelObjectHandle
    {
        public object Object { get; }

        internal object[] CallerParameters { get; }

        public ExcelObjectHandle(object o, object[] callerParameters)
        {
            Object = o;
            CallerParameters = callerParameters;
        }
    }
}
