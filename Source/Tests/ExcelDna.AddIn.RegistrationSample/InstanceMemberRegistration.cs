using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Registration;

namespace ExcelDna.AddIn.RegistrationSample
{
    // Test class
    [AttributeUsage(AttributeTargets.All)]
    public class ExcelObjectAttribute : Attribute
    {
    }

    [ExcelObject]
    public class TestClass
    {
        string _content;

        public TestClass(string content)
        {
            _content = content;
        }

        public string GetContent()
        {
            return "Content is " + _content;
        }

        // Not done yet...
        //public string Content
        //{
        //    get { return _content; }
        //}
    }

    class InstanceMemberRegistration
    {
        public static LambdaExpression WrapInstanceMethod(MethodInfo method)
        {
            // We wrap a method in class MyType from
            //      public object MyFunction(object arg1, object arg2) {...}
            // with a lambda expression that looks like this
            //      (MyType myType, object arg1, object arg2) => instance(arg1, arg2)

            var instanceParam = Expression.Parameter(method.DeclaringType, method.DeclaringType.Name);
            var callParams = method.GetParameters().Select(p => Expression.Parameter(p.ParameterType, p.Name)).ToList();
            var allParams = new List<ParameterExpression>();
            allParams.Add(instanceParam);
            allParams.AddRange(callParams);

            var callExpr = Expression.Call(instanceParam, method, callParams);
            return Expression.Lambda(callExpr, method.Name, allParams);
        }

        public static void TestInstanceRegistration()
        {
            var instanceMethod = typeof(TestClass).GetMethod("GetContent");
            var lambda = WrapInstanceMethod(instanceMethod);

            // Create a new object every time the instance method is called... (typically we'd look up using a handle)
            var paramConversionConfig = new ParameterConversionConfiguration()
                .AddParameterConversion((string content) => new TestClass(content));

            // TODO: Clean up how we suggest this code gets called
            var reg = new ExcelFunctionRegistration(lambda);
            var processed = ParameterConversionRegistration.ProcessParameterConversions(new[] { reg }, paramConversionConfig);
            processed.RegisterFunctions();
        }
    }
}
