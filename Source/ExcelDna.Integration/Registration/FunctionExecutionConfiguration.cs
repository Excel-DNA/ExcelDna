using System;
using System.Collections.Generic;

namespace ExcelDna.Registration
{
    public class FunctionExecutionConfiguration
    {
        internal List<Func<ExcelFunctionRegistration, IFunctionExecutionHandler>> FunctionHandlerSelectors { get; private set; }

        public FunctionExecutionConfiguration()
        {
            FunctionHandlerSelectors = new List<Func<ExcelFunctionRegistration, IFunctionExecutionHandler>>();
        }

        public FunctionExecutionConfiguration AddFunctionExecutionHandler(Func<ExcelFunctionRegistration, IFunctionExecutionHandler> functionHandlerSelector)
        {
            FunctionHandlerSelectors.Add(functionHandlerSelector);
            return this;
        }
    }
}
