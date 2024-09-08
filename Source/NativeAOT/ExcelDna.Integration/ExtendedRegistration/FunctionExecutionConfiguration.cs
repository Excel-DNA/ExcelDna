using System;
using System.Collections.Generic;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal class FunctionExecutionConfiguration
    {
        internal List<Func<ExcelFunction, IFunctionExecutionHandler>> FunctionHandlerSelectors { get; private set; }

        public FunctionExecutionConfiguration()
        {
            FunctionHandlerSelectors = new List<Func<ExcelFunction, IFunctionExecutionHandler>>();
        }

        public FunctionExecutionConfiguration AddFunctionExecutionHandler(Func<ExcelFunction, IFunctionExecutionHandler> functionHandlerSelector)
        {
            FunctionHandlerSelectors.Add(functionHandlerSelector);
            return this;
        }
    }
}
