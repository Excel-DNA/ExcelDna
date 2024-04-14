using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using Expr = System.Linq.Expressions.Expression;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal static class FunctionExecutionRegistration
    {
        public static IEnumerable<ExcelFunction> ProcessFunctionExecutionHandlers(this IEnumerable<ExcelFunction> registrations, FunctionExecutionConfiguration functionHandlerConfig)
        {
            foreach (var registration in registrations)
            {
                var reg = registration; // Ensure safe semantics for captured foreach variable

                // Exclude the functions created for native async, with no return values.
                // Can't deal with these yet.
                if (reg.FunctionLambda.ReturnType != typeof(void))
                {
                    var handlers = functionHandlerConfig.FunctionHandlerSelectors
                                                      .Select(fhSelector => fhSelector(reg))
                                                      .Where(fh => fh != null);
                    ApplyMethodHandlers(reg, handlers);
                }
                yield return reg;
            }
        }

        static void ApplyMethodHandlers(ExcelFunction reg, IEnumerable<IFunctionExecutionHandler> handlers)
        {
            // The order of method handlers is important - we follow PostSharp's convention for MethodExecutionHandlers.
            // They are passed from high priority (most inside) to low priority (most outside)
            // Imagine 2 FunctionHandlers, fh1 then fh2
            // So fh1 (highest priority)  will be 'inside' and fh2 will be outside (lower priority)
            foreach (var handler in handlers)
            {
                reg.FunctionLambda = ApplyMethodHandler(reg.FunctionAttribute.Name, reg.FunctionLambda, handler);
            }
        }

        static LambdaExpression ApplyMethodHandler(string functionName, LambdaExpression functionLambda, IFunctionExecutionHandler handler)
        {
            // public static int MyMethod(object arg0, int arg1) { ... }

            // becomes:

            // (the 'handler' object is captured and called mh)
            // public static int MyMethodWrapped(object arg0, int arg1)
            // {
            //    var fhArgs = new FunctionExecutionArgs("MyMethod", new object[] { arg0, arg1});
            //    int result = default(int);
            //    try
            //    {
            //        fh.OnEntry(fhArgs);
            //        if (fhArgs.FlowBehavior == FlowBehavior.Return)
            //        {
            //            result = (int)fhArgs.ReturnValue;
            //        }
            //        else
            //        {
            //             // Inner call
            //             result = MyMethod(arg0, arg1);
            //             fhArgs.ReturnValue = result;
            //             fh.OnSuccess(fhArgs);
            //             result = (int)fhArgs.ReturnValue;
            //        }
            //    }
            //    catch ( Exception ex )
            //    {
            //        fhArgs.Exception = ex;
            //        fh.OnException(fhArgs);
            //        // Makes no sense to me yet - I've removed this FlowBehavior enum value.
            //        // if (fhArgs.FlowBehavior == FlowBehavior.Continue)
            //        // {
            //        //     // Finally will run, but can't change return value
            //        //     // Should we assign result...?
            //        //     // So Default value will be returned....?????
            //        //     fhArgs.Exception = null;
            //        // }
            //        // else 
            //        if (fhArgs.FlowBehavior == FlowBehavior.Return)
            //        {
            //            // Clear the Exception and return the ReturnValue instead
            //            // Finally will run, but can't further change return value
            //            fhArgs.Exception = null;
            //            result = (int)fhArgs.ReturnValue;
            //        }
            //        else if (fhArgs.FlowBehavior == FlowBehavior.ThrowException)
            //        {
            //            throw fhArgs.Exception;   // TODO: Check if we can capture and add context to the throw here
            //        }
            //        else // if (fhArgs.FlowBehavior == FlowBehavior.Default || fhArgs.FlowBehavior == FlowBehavior.RethrowException)
            //        {
            //            throw;
            //        }
            //    }
            //    finally
            //    {
            //        fh.OnExit(fhArgs);
            //        // NOTE: fhArgs.ReturnValue is not used again here...!
            //    }
            //    
            //    return result;
            //  }
            // }

            // CONSIDER: There are some helpers in .NET to capture the exception context, which would allow us to preserve the stack trace in a fresh throw.

            // Ensure the handler object is captured.
            var mh = Expression.Constant(handler);
            var funcName = Expression.Constant(functionName);

            // Prepare the functionHandlerArgs that will be threaded through the handler, 
            // and a bunch of expressions that access various properties on it.
            var fhArgs = Expr.Variable(typeof(FunctionExecutionArgs), "fhArgs");
            var fhArgsReturnValue = SymbolExtensions.GetProperty(fhArgs, (FunctionExecutionArgs mea) => mea.ReturnValue);
            var fhArgsException = SymbolExtensions.GetProperty(fhArgs, (FunctionExecutionArgs mea) => mea.Exception);
            var fhArgsFlowBehaviour = SymbolExtensions.GetProperty(fhArgs, (FunctionExecutionArgs mea) => mea.FlowBehavior);

            // Set up expressions to call the various handler methods.
            // TODO: Later we can determine which of these are actually implemented, and only write out the code needed in the particular case.
            var onEntry = Expr.Call(mh, SymbolExtensions.GetMethodInfo<IFunctionExecutionHandler>(meh => meh.OnEntry(null)), fhArgs);
            var onSuccess = Expr.Call(mh, SymbolExtensions.GetMethodInfo<IFunctionExecutionHandler>(meh => meh.OnSuccess(null)), fhArgs);
            var onException = Expr.Call(mh, SymbolExtensions.GetMethodInfo<IFunctionExecutionHandler>(meh => meh.OnException(null)), fhArgs);
            var onExit = Expr.Call(mh, SymbolExtensions.GetMethodInfo<IFunctionExecutionHandler>(meh => meh.OnExit(null)), fhArgs);

            // Create the new parameters for the wrapper
            var outerParams = functionLambda.Parameters.Select(p => Expr.Parameter(p.Type, p.Name)).ToArray();
            // Create the array of parameter values that will be put into the method handler args.
            var paramsArray = Expr.NewArrayInit(typeof(object), outerParams.Select(p => Expr.Convert(p, typeof(object))));

            // Prepare the result and ex(ception) local variables
            var result = Expr.Variable(functionLambda.ReturnType, "result");
            var ex = Expression.Parameter(typeof(Exception), "ex");

            // A bunch of helper expressions:
            // : new FunctionExecutionArgs(new object[] { arg0, arg1 })
            var fhArgsConstr = typeof(FunctionExecutionArgs).GetConstructor(new[] { typeof(string), typeof(object[]) });
            var newfhArgs = Expr.New(fhArgsConstr, funcName, paramsArray);
            // : result = (int)fhArgs.ReturnValue
            var resultFromReturnValue = Expr.Assign(result, Expr.Convert(fhArgsReturnValue, functionLambda.ReturnType));
            // : fhArgs.ReturnValue = (object)result
            var returnValueFromResult = Expr.Assign(fhArgsReturnValue, Expr.Convert(result, typeof(object)));
            // : result = function(arg0, arg1)
            var resultFromInnerCall = Expr.Assign(result, Expr.Invoke(functionLambda, outerParams));

            // Build the Lambda wrapper, with the original parameters
            var lambda = Expr.Lambda(
                Expr.Block(new[] { fhArgs, result },
                     Expr.Assign(fhArgs, newfhArgs),
                     Expr.Assign(result, Expr.Default(result.Type)),
                     Expr.TryCatchFinally(
                        Expr.Block(
                            onEntry,
                            Expr.IfThenElse(
                                Expr.Equal(fhArgsFlowBehaviour, Expr.Constant(FlowBehavior.Return)),
                                resultFromReturnValue,
                                Expr.Block(
                                    resultFromInnerCall,
                                    returnValueFromResult,
                                    onSuccess,
                                    resultFromReturnValue))),
                        onExit, // finally
                        Expr.Catch(ex,
                            Expr.Block(
                                Expr.Assign(fhArgsException, ex),
                                onException,
                                Expr.IfThenElse(
                                    Expr.Equal(fhArgsFlowBehaviour, Expr.Constant(FlowBehavior.Return)),
                                    Expr.Block(
                                        Expr.Assign(fhArgsException, Expr.Constant(null, typeof(Exception))),
                                        resultFromReturnValue),
                                    Expr.IfThenElse(
                                        Expr.Equal(fhArgsFlowBehaviour, Expr.Constant(FlowBehavior.ThrowException)),
                                        Expr.Throw(fhArgsException),
                                        Expr.Rethrow()))))
                        ),
                    result),
                functionName,
                outerParams);
            return lambda;
        }
    }
}
