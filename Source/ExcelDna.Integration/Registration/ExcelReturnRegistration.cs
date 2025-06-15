using ExcelDna.Integration;
using System.Collections.Generic;

namespace ExcelDna.Registration
{
    public class ExcelReturnRegistration : IExcelFunctionReturn
    {
        // Used only for the Registration processing
        public List<object> CustomAttributes { get; private set; } // Should not be null, and elements should not be null

        public ExcelReturnRegistration()
        {
            CustomAttributes = new List<object>();
        }
    }
}
