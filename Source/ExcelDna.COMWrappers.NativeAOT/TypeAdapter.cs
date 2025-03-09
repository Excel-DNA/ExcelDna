using Addin.ComApi;
using ExcelDna.ComInterop;

namespace ExcelDna.COMWrappers.NativeAOT
{
    public class TypeAdapter : IType
    {
        public object GetProperty(string name, object comObject)
        {
            var excelWindowWrapper = new ExcelObject(comObject as IDispatch);
            object? result = excelWindowWrapper.GetProperty(name);
            if (result is ExcelObject excelObject)
                return excelObject._interfacePtr;

            return result!;
        }

        public object GetIndex(int i, object comObject)
        {
            var excelWindowWrapper = new ExcelObject(comObject as IDispatch);
            excelWindowWrapper.TryGetIndex(null, new object[] { i }, out object? result);
            if (result is ExcelObject excelObject)
                return excelObject._interfacePtr;

            return result!;
        }

        public bool Is(ref Guid guid, object comObject)
        {
            if (guid == new Guid("000C030E-0000-0000-C000-000000000046"))
                return comObject is ICommandBarButton;

            if (guid == new Guid("000C030A-0000-0000-C000-000000000046"))
                return comObject is ICommandBarPopup;

            if (guid == new Guid("000C030C-0000-0000-C000-000000000046"))
                return comObject is ICommandBarComboBox;

            throw new NotImplementedException();
        }

        public object Invoke(string name, object[] args, object comObject)
        {
            var excelWindowWrapper = new ExcelObject(comObject as IDispatch);
            object? result = excelWindowWrapper.InvokeMember(name, args);
            if (result is ExcelObject excelObject)
                return excelObject._interfacePtr;

            return result!;
        }
    }
}
