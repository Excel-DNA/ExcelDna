using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using ExcelDna.Integration;
using NUnit.Framework;

namespace ExcelDna.Registration.Test
{
    [TestFixture]
    public static class ParameterConversionTests
    {
        [Test]
        public static void TestReturnConversion()
        {
            // 
            var testLambda = (Expression<Func<string, int, string>>)((name, value) => name + value.ToString());
            var registration = new ExcelFunctionRegistration(testLambda);

            // This conversion takes the return string and duplicates it
            var conversionConfig = new ParameterConversionConfiguration()
                                    .AddReturnConversion((string value) => value + value);

            var convertedRegistrations = ParameterConversionRegistration.ProcessParameterConversions(new[] { registration }, conversionConfig);
            var converted = convertedRegistrations.First();
            var compiled = (Func<string, int, string>)converted.FunctionLambda.Compile();
            var result = compiled("asd", 42);
            Assert.AreEqual("asd42asd42", result);
        }

        [Test]
        public static void TestParameterConversions()
        {
            // 
            var testLambda = (Expression<Func<string, int, string>>)((name, value) => name + value.ToString());
            var registration = new ExcelFunctionRegistration(testLambda);

            // This conversion applied a mix of parameter and string conversions
            var conversionConfig = new ParameterConversionConfiguration()
                                    .AddParameterConversion((double value) => (int)(value * 10))
                                    .AddParameterConversion((string value) => value.Substring(0, 2))
                                    .AddReturnConversion((string value) => value + value)
                                    .AddReturnConversion((string value) => value.Length);

            var convertedRegistrations = ParameterConversionRegistration.ProcessParameterConversions(new[] { registration }, conversionConfig);
            var converted = convertedRegistrations.First();
            var compiled = (Func<string, double, int>)converted.FunctionLambda.Compile();
            var result = compiled("qwXXX", 4.2);
            Assert.AreEqual("qw42qw42".Length, result);
        }

        [Test]
        public static void TestGlobalReturnConversion()
        {
            // Conversions can be 'global' so they can be applied to any function.
            // The conversion returns 'null' for cases (type + attribute combinations) where it should not be applied.

            // This test simulates the return type conversion we would do to change async #N/A to something else
            // (though using strings instead of the real ExcelError.ExcelErrorNA result).

            var testLambda = (Expression<Func<string, object>>)(name => name);
            var registration = new ExcelFunctionRegistration(testLambda);

            var conversionConfig = new ParameterConversionConfiguration()
                .AddReturnConversion((type, customAttributes) => type != typeof(object) ? null : ((Expression<Func<object, object>>)
                                                ((object returnValue) => returnValue.Equals("NA") ? (object)"### WAIT ###" : returnValue)), null);

            var convertedRegistrations = ParameterConversionRegistration.ProcessParameterConversions(new[] { registration }, conversionConfig);
            var converted = convertedRegistrations.First();
            var compiled = (Func<string, object>)converted.FunctionLambda.Compile();
            
            var resultNormal = compiled("XYZ");
            Assert.AreEqual("XYZ", resultNormal);

            var resultNA = compiled("NA");
            Assert.AreEqual("### WAIT ###", resultNA);
        }

        static object TestMethodWithLambda(string value)
        {
            var doubleFunc = (Func<string>)(() => value + value);
            return doubleFunc();
        }

        [Test]
        public static void TestNestedLambdaConversion()
        {
            // The function duplicates the input string (inside a lambda that captures that input parameter)
            var registration = new ExcelFunctionRegistration(typeof(ParameterConversionTests).GetMethod("TestMethodWithLambda", BindingFlags.NonPublic | BindingFlags.Static));

            var conversionConfig = new ParameterConversionConfiguration()
                .AddReturnConversion((type, customAttributes) => type != typeof(object) ? null : ((Expression<Func<object, object>>)
                                                ((object returnValue) => returnValue.Equals("NANA") ? (object)"### WAIT ###" : returnValue)), null);

            var convertedRegistrations = ParameterConversionRegistration.ProcessParameterConversions(new[] { registration }, conversionConfig);
            var converted = convertedRegistrations.First();
            var compiled = (Func<string, object>)converted.FunctionLambda.Compile();

            var resultNormal = compiled("XYZ");
            Assert.AreEqual("XYZXYZ", resultNormal);

            var resultNA = compiled("NA");
            Assert.AreEqual("### WAIT ###", resultNA);
        }

        [Test]
        public static void TestIgnoredGlobalReturnConversion()
        {
            var testLambda = (Expression<Func<string, object>>)(name => name);
            var registration = new ExcelFunctionRegistration(testLambda);

            var conversionConfig = new ParameterConversionConfiguration()
                // Add a return conversion that is never applied
                .AddReturnConversion((type, customAttributes) => null, null);

            var convertedRegistrations = ParameterConversionRegistration.ProcessParameterConversions(new[] { registration }, conversionConfig);
            var converted = convertedRegistrations.First();
            var compiled = (Func<string, object>)converted.FunctionLambda.Compile();

            var resultNormal = compiled("XYZ");
            Assert.AreEqual("XYZ", resultNormal);
        }

    }
}
