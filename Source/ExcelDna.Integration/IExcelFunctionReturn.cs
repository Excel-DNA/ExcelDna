using System.Collections.Generic;

namespace ExcelDna.Integration
{
    public interface IExcelFunctionReturn
    {
        List<object> CustomAttributes { get; }
    }
}
