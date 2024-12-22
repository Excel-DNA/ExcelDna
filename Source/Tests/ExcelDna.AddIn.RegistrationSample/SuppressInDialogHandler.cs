using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.Registration;

namespace ExcelDna.AddIn.RegistrationSample
{
    // Alternative pattern - make this an attribute directly
    // CONSIDER: In this case we never need the parameters. It would be nice never have to pull them into the FunctionExecutionArgs.

    // TODO: Only works for functions that return string or object. Automatically add a return value conversion otherwise?

    [AttributeUsage(AttributeTargets.Method)]
    public class SuppressInDialogAttribute : Attribute, IFunctionExecutionHandler
    {
        readonly string _dialogResult;
        public SuppressInDialogAttribute(string dialogMessage = "!!! NOT CALCULATED IN DIALOG !!!")
        {
            _dialogResult = dialogMessage;
        }

        public void OnEntry(FunctionExecutionArgs args)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            {
                args.ReturnValue = _dialogResult;
                args.FlowBehavior = FlowBehavior.Return;
            }
            // Otherwise we do not interfere
        }

        // Implemented just to satisfy the interface
        public void OnSuccess(FunctionExecutionArgs args) { }
        public void OnException(FunctionExecutionArgs args) { }
        public void OnExit(FunctionExecutionArgs args) { }
    }

    public static class SuppressInDialogFunctionExecutionHandler
    {
        /// <summary>
        /// Currently only applied to functions that return object or string.
        /// </summary>
        /// <param name="functionRegistration"></param>
        /// <returns></returns>
        public static IFunctionExecutionHandler SuppressInDialogSelector(ExcelFunctionRegistration functionRegistration)
        {
            // Eat the TimingAttributes, and return a timer handler if there were any
            if (functionRegistration.CustomAttributes.OfType<SuppressInDialogAttribute>().Any() &&
                (functionRegistration.FunctionLambda.ReturnType == typeof(object) ||
                 functionRegistration.FunctionLambda.ReturnType == typeof(string)))
            {
                // Get the first cache attribute, and remove all of them
                var suppressAtt = functionRegistration.CustomAttributes.OfType<SuppressInDialogAttribute>().First();
                functionRegistration.CustomAttributes.RemoveAll(att => att is SuppressInDialogAttribute);

                return suppressAtt;
            }
            return null;

        }

    }
}
