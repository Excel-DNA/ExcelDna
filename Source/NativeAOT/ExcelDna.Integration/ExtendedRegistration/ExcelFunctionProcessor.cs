using System.Collections.Generic;
using System.Reflection;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal class ExcelFunctionProcessor
    {
        private MethodInfo mi;

        public string Name => mi.Name;

        public ExcelFunctionProcessor(MethodInfo mi)
        {
            this.mi = mi;
        }

        public IEnumerable<IExcelFunctionInfo> Invoke(IEnumerable<IExcelFunctionInfo> registrations, IExcelFunctionRegistrationConfiguration config)
        {
            return (IEnumerable<IExcelFunctionInfo>)mi.Invoke(null, new object[] { registrations, config });
        }
    }
}
