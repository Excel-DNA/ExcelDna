using ExcelDna.Integration;
using ExcelDna.Registration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDna.AddIn.RegistrationSampleRuntimeTests
{
    public class ExampleAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex => "!!! ERROR: " + ex.ToString());

            ExcelRegistration.GetExcelFunctions().RegisterFunctions();
        }

        public void AutoClose()
        {
        }
    }
}
