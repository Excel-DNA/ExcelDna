using System.Collections.Generic;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal class ExcelReturn : IExcelFunctionReturn
    {
        // Used only for the Registration processing
        public List<object> CustomAttributes { get; private set; } // Should not be null, and elements should not be null

        public ExcelReturn()
        {
            CustomAttributes = new List<object>();
        }
    }
}
