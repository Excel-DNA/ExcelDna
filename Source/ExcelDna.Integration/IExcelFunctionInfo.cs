using System.Collections.Generic;

namespace ExcelDna.Integration
{
    public interface IExcelFunctionInfo
    {
        List<object> CustomAttributes { get; }
    }
}
