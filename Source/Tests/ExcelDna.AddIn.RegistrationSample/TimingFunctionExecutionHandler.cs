using System;
using System.Diagnostics;
using ExcelDna.Registration;

namespace ExcelDna.AddIn.RegistrationSample
{
    [AttributeUsage(AttributeTargets.Method)]
    public class TimingAttribute : Attribute
    {
    }

    public class TimingFunctionExecutionHandler : FunctionExecutionHandler
    {
        public override void OnEntry(FunctionExecutionArgs args)
        {
            args.Tag = Stopwatch.StartNew();
        }

        public override void OnExit(FunctionExecutionArgs args)
        {
            var sw = (Stopwatch)args.Tag;
            sw.Stop();

            Debug.WriteLine("TimingFunctionExecutionHandler {0} executed in {1} milliseconds", args.FunctionName, sw.ElapsedMilliseconds);
            Logger.Log($"TimingFunctionExecutionHandler {args.FunctionName}");
        }

        /////////////////////// Registration handler //////////////////////////////////

        // In this case, we only ever make one 'handler' object
        static readonly Lazy<TimingFunctionExecutionHandler> _handler =
            new Lazy<TimingFunctionExecutionHandler>(() => new TimingFunctionExecutionHandler());

        internal static FunctionExecutionHandler TimingHandlerSelector(ExcelFunctionRegistration functionRegistration)
        {
            // Eat the TimingAttributes, and return a timer handler if there were any
            if (functionRegistration.CustomAttributes.RemoveAll(att => att is TimingAttribute) == 0)
            {
                // No attributes
                return null;
            }
            return _handler.Value;
        }
    }
}
