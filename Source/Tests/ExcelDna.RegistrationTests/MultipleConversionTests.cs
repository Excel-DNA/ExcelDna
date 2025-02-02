using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using NUnit.Framework;
using Expr = System.Linq.Expressions.Expression;

namespace ExcelDna.Registration.Test
{
    [TestFixture]
    public static class MultipleConversionTests
    {

        //[Test]
        //public static void TestAsyncAndReturnConversion()
        //{
        //    // 
        //    var testLambda = (Expression<Func<string, Task<string>>>)(arg => Task.FromResult(arg));

        //    var testCompiled = testLambda.Compile();
        //    var result1 = testCompiled("xyz");
        //    Assert.AreEqual("xyz", result1.Result);
        //    var wrapped = WrapTaskMethod(testLambda);
        //    //var registration = new ExcelFunctionRegistration(wrapped);
        //    // var registration = new ExcelFunctionRegistration(testLambda);

        //    // This conversion takes the return string and duplicates it
        //    //var conversionConfig = new ParameterConversionConfiguration()
        //    //.AddReturnConversion((string value) => value + value)
        //    //;

        //    //va//r convertedRegistrations = ParameterConversionRegistration.ProcessParameterConversions(new[] { registration }, conversionConfig);
        //    //var converted = convertedRegistrations.First();
        //    var converted = EmptyWrapper(wrapped);
        //    var compiled = (Func<string, object>)converted.Compile();
        //    //var compiled = (Func<string, object>)converted.Compile();
        //    var result = compiled("asd");
        //    //            Assert.AreEqual("asd", result);
        //    Assert.AreEqual("asd", ((Task<string>)result).Result);
        //}


        [Test]
        public static void TestAsyncAndReturnConversion()
        {
            // 
            var testLambda = (Expression<Func<string, string>>)(arg => arg);

            var testCompiled = testLambda.Compile();
            var result1 = testCompiled("xyz");
            Assert.AreEqual("xyz", result1);
            var wrapped = WrapMethod(testLambda);
            var wrappedCompiled = (Func<string, string>)wrapped.Compile();
            var result2 = wrappedCompiled("xyz");
            //wrapped = WrapMethod(wrapped);
            //wrapped = testLambda;
            //var registration = new ExcelFunctionRegistration(wrapped);
            // var registration = new ExcelFunctionRegistration(testLambda);

            // This conversion takes the return string and duplicates it
            //var conversionConfig = new ParameterConversionConfiguration()
                                    //.AddReturnConversion((string value) => value + value)
                                    //;

            //va//r convertedRegistrations = ParameterConversionRegistration.ProcessParameterConversions(new[] { registration }, conversionConfig);
            //var converted = convertedRegistrations.First();
            var converted = EmptyWrapper(wrapped);
            //converted = EmptyWrapper(converted);
            //var converted = wrapped;
            var compiled = (Func<string, string>)converted.Compile();
            //var compiled = (Func<string, object>)converted.Compile();
            var result = compiled("asd");
            Assert.AreEqual("asd", result);
        }

        // This is a helper that is similar to the async wrapper, but using a local version of the Task helper
        // which just sets the result and returns.
        static LambdaExpression WrapMethod(LambdaExpression functionLambda)
        {
            /* From a lambda expression wrapping a method that looks like this:
             * 
             *      static Task<string> myFunc(string name, int msDelay) {...}
             * 
             *   we create a lambda expression that looks like this:
             * 
             *      static object myFunc(string nameX, int msDelayX)
             *      {
             *          return MultipleConversionTests.ReturnTaskResult<string>(
             *              "myFunc:XXX", 
             *              new object[] {(object)nameX, (object)msDelayX}, 
             *              () => myFunc(nameX, msDelayX));
             *      }
             */

            //string runMethodName = "ReturnResult";
            // mi returns some kind of Task<T>. What is T? 
            // Build up the RunTaskWithC... method with the right generic type argument
            //var runMethod = typeof(MultipleConversionTests)
            //                    .GetMember(runMethodName, MemberTypes.Method, BindingFlags.Static | BindingFlags.Public)
            //                    .Cast<MethodInfo>().First();

            //// Get the function name
            //var nameExp = Expression.Constant(functionLambda.Name + ":" + Guid.NewGuid().ToString("N"));

            var newParams = functionLambda.Parameters.Select(p => Expression.Parameter(p.Type, p.Name)).ToList();

            //// Also cast params to Object and put into a fresh object[] array for the RunTask call
            //var paramsArray = functionLambda.Parameters.Select(p => Expression.Convert(p, typeof(object)));
            //var paramsArrayExp = Expression.NewArrayInit(typeof(object), paramsArray);
            var innerLambda = Expression.Lambda(Expression.Invoke(functionLambda, functionLambda.Parameters));

            // This is the call to RunTask, taking the name, param array and the (capturing) lambda (called with no arguments)
//            var callTaskRun = Expression.Call(runMethod, nameExp, paramsArrayExp, innerLambda);
            //var callRun = Expression.Call(runMethod, innerLambda);
            Expression<Func<Func<string>, string>> callLambda = f => f();
            var callRun = Expression.Invoke(callLambda, innerLambda);
            // Wrap with all the parameters, and Compile to a Delegate
            var lambda = Expression.Lambda(callRun, functionLambda.Parameters);
            return lambda;
        }

        static LambdaExpression EmptyWrapper(LambdaExpression functionLambda)
        {
            // We do the following transformation
            //     arg => arg
            // to
            //    arg => (arg => arg)(arg)
            //      public static string dnaParameterConvertTest(string optTest) {  ... };
            //
            // to
            //      (optTest) => dnaParameterConvertTest(optTest)
            ////      public static string dnaParameterConvertTest(string optTest) 
            ////      {   
            ////          return dnaParameterConvertTest(optTest);
            ////      };

            var newParams = functionLambda.Parameters.Select(p => Expression.Parameter(p.Type, p.Name)).ToList();
            
            var call = Expr.Invoke(functionLambda, newParams);
            return Expr.Lambda(call, functionLambda.Name, newParams);


            //// build up the invoke expression for each parameter
            //var wrappingParameters = new List<ParameterExpression>(functionLambda.Parameters);
            //var paramExprs = functionLambda.Parameters.Select((param, i) =>
            //{
            //    // Starting point is just the parameter expression
            //    Expression wrappedExpr = param;
            //    return wrappedExpr;
            //}).ToArray();

            //var wrappingCall = Expr.Invoke(functionLambda, paramExprs);
            //return Expr.Lambda(wrappingCall, functionLambda.Name, wrappingParameters);
        }

        public static string ReturnResult(Func<string> stringFactory)
        {
            var result = stringFactory();
            return result;
        }
    }
}
