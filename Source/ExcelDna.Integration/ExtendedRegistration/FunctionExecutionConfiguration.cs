using System;
using System.Collections.Generic;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal class FunctionExecutionConfiguration
    {
        internal List<Func<ExcelDna.Registration.ExcelFunctionRegistration, IFunctionExecutionHandler>> FunctionHandlerSelectors { get; private set; }

        public FunctionExecutionConfiguration()
        {
            FunctionHandlerSelectors = new List<Func<ExcelDna.Registration.ExcelFunctionRegistration, IFunctionExecutionHandler>>();
        }

        public FunctionExecutionConfiguration AddFunctionExecutionHandler(Func<ExcelDna.Registration.ExcelFunctionRegistration, IFunctionExecutionHandler> functionHandlerSelector)
        {
            FunctionHandlerSelectors.Add(functionHandlerSelector);
            return this;
        }
    }
}
