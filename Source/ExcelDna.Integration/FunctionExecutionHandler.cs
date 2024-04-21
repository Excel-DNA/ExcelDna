using System;

namespace ExcelDna.Integration
{
    // Inspired by PostSharp
    public enum FlowBehavior
    {
        /// <summary>
        /// Default behaviour - Same as continue for OnEnter, OnSuccess and OnExit; same as RethrowException for OnException.
        /// </summary>
        Default = 0,
        // Makes no sense to me yet.
        ///// <summary>
        ///// Continue normally - For an OnException handler would suppress the exception and continue as if the method were successful
        ///// </summary>
        //Continue = 1,
        /// <summary>
        /// Rethrow the current exception - only valid for OnException handlers.
        /// </summary>
        RethrowException = 2,
        /// Return the value of ReturnValue immediately  - For OnEnter will skip the method execution and the OnSuccess handlers, but will run OnExit handlers
        Return = 3,
        /// <summary>
        /// Throw the Exception in the Exception property - For OnException handlers only.
        /// </summary>
        ThrowException = 4
    }

    // CONSIDER: One might make a generic typed version of this...
    public class FunctionExecutionArgs
    {
        public string FunctionName { get; set; }
        // Can't change arguments - Make ReadOnly collection?
        public object[] Arguments { get; private set; }
        public object ReturnValue { get; set; }
        public Exception Exception { get; set; }
        public FlowBehavior FlowBehavior { get; set; }
        public object Tag { get; set; }

        public FunctionExecutionArgs(string functionName, object[] arguments)
        {
            FunctionName = functionName;
            Arguments = arguments;
        }
    }

    /*
        // Conceptually we rewrite as 
        public static int MyMethodWrapped(object arg0, int arg1)
        {
          int result;
          try
          {
            OnEntry();
            result = MyMethod(arg0, arg1);
            OnSuccess();
          }
          catch ( Exception e )
          {
            OnException();
          }
          finally
          {
            OnExit();
          }
          return result;
        }
     * 
     *  However there are advanced options to understand too...
    */

    public interface IFunctionExecutionHandler
    {
        void OnEntry(FunctionExecutionArgs args);
        void OnSuccess(FunctionExecutionArgs args);
        void OnException(FunctionExecutionArgs args);
        void OnExit(FunctionExecutionArgs args);
    }

    // Can inherit from here or implement interface directly
    public class FunctionExecutionHandler : IFunctionExecutionHandler
    {
        public virtual void OnEntry(FunctionExecutionArgs args) { }
        public virtual void OnSuccess(FunctionExecutionArgs args) { }
        public virtual void OnException(FunctionExecutionArgs args) { }
        public virtual void OnExit(FunctionExecutionArgs args) { }
    }

    public delegate IFunctionExecutionHandler FunctionExecutionHandlerSelector(IExcelFunctionInfo functionInfo);
}
