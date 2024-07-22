using System.Collections.Generic;

namespace ExcelDna.Integration
{
    public interface IExcelFunctionParameter
    {
        ExcelArgumentAttribute ArgumentAttribute { get; }
        List<object> CustomAttributes { get; }
    }
}
