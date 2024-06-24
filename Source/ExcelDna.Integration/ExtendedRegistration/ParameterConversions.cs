using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelDna.Integration.ExtendedRegistration
{
    /// <summary>
    /// Defines some standard Parameter Conversions.
    /// Register by calling ParameterConversionConfiguration.AddParameterConversion(ParameterConversions.NullableConversion);
    /// </summary>
    internal static class ParameterConversions
    {
        // These can be used directly in .AddParameterConversion

        /// <summary>
        /// Legacy method: this returns a converter for Nullable[T] where T is one of the basic types that do not require any converter.
        /// If you need a Nullable[T] converter that can call into another for T, then use ParameterConversionConfiguration.AddNullableConversion.
        /// </summary>
        /// <param name="treatEmptyAsMissing"></param>
        /// <param name="treatNAErrorAsMissing"></param>
        /// <returns></returns>
        public static Func<Type, ExcelParameter, LambdaExpression> GetNullableConversion(bool treatEmptyAsMissing = false, bool treatNAErrorAsMissing = false)
        {
            return (type, paramReg) => NullableConversion(null, type, paramReg, treatEmptyAsMissing, treatNAErrorAsMissing);
        }

        public static Func<Type, ExcelParameter, LambdaExpression> GetOptionalConversion(bool treatEmptyAsMissing = false, bool treatNAErrorAsMissing = false)
        {
            return (type, paramReg) => OptionalConversion(type, paramReg, treatEmptyAsMissing, treatNAErrorAsMissing);
        }

        public static Func<Type, ExcelParameter, LambdaExpression> GetEnumStringConversion()
        {
            return (type, paramReg) => EnumStringConversion(type, paramReg);
        }

        public static IEnumerable<Func<Type, ExcelParameter, LambdaExpression>> GetUserConversions(IEnumerable<ExcelParameterConversion> parameterConversions)
        {
            return parameterConversions.OrderBy(i => i.MethodInfo.Name).Select(i => i.GetConversion());
        }

        internal static LambdaExpression NullableConversion(
            ParameterConversionConfiguration config, Type type,
            ExcelParameter paramReg, bool treatEmptyAsMissing,
            bool treatNAErrorAsMissing)
        {
            // Decide whether to return a conversion function for this parameter
            if (!type.IsGenericType || type.GetGenericTypeDefinition() != typeof(Nullable<>))
                return null;

            var innerType = type.GetGenericArguments()[0]; // E.g. innerType is Complex
            LambdaExpression innerTypeConversion = ParameterConversionRegistration.GetParameterConversion(config, innerType, paramReg) ??
                                                   TypeConversion.GetConversion(typeof(object), innerType);
            ParameterExpression input = innerTypeConversion.Parameters[0];
            // Here's the actual conversion function
            var result =
                Expression.Lambda(
                    Expression.Condition(
                        // if the value is missing (or possibly empty)
                        MissingTest(input, treatEmptyAsMissing, treatNAErrorAsMissing),
                        // cast null to int?
                        Expression.Constant(null, type),
                        // else convert to int, and cast that to int?
                        Expression.Convert(Expression.Invoke(innerTypeConversion, input), type)),
                    input);
            return result;
        }

        static LambdaExpression OptionalConversion(Type type, ExcelParameter paramReg, bool treatEmptyAsMissing, bool treatNAErrorAsMissing)
        {
            // Decide whether to return a conversion function for this parameter
            if (!paramReg.CustomAttributes.OfType<OptionalAttribute>().Any())
                return null;

            var defaultAttribute = paramReg.CustomAttributes.OfType<DefaultParameterValueAttribute>().FirstOrDefault();
            var defaultValue = defaultAttribute == null ? TypeConversion.GetDefault(type) : defaultAttribute.Value;
            // var returnType = type.GetGenericArguments()[0]; // E.g. returnType is double

            // Consume the attributes
            paramReg.CustomAttributes.RemoveAll(att => att is OptionalAttribute);
            paramReg.CustomAttributes.RemoveAll(att => att is DefaultParameterValueAttribute);

            // Here's the actual conversion function
            var input = Expression.Parameter(typeof(object), "input");
            return
                Expression.Lambda(
                    Expression.Condition(
                        MissingTest(input, treatEmptyAsMissing, treatNAErrorAsMissing),
                        Expression.Constant(defaultValue, type),
                        TypeConversion.GetConversion(input, type)),
                    input);
        }

        static Expression MissingTest(ParameterExpression input, bool treatEmptyAsMissing,
            bool treatNAErrorAsMissing)
        {
            Expression r = null;
            if (treatNAErrorAsMissing)
            {
                var methodMissingOrNATest = typeof(ParameterConversions).GetMethod("MissingOrNATest",
                    BindingFlags.NonPublic | BindingFlags.Static);
                r = Expression.Call(null, methodMissingOrNATest, input, Expression.Constant(treatEmptyAsMissing));
            }
            else
            {
                r = Expression.TypeIs(input, typeof(ExcelMissing));
                if (treatEmptyAsMissing)
                    r = Expression.OrElse(r, Expression.TypeIs(input, typeof(ExcelEmpty)));
            }
            return r;
        }

        static bool MissingOrNATest(object input, bool treatEmptyAsMissing)
        {
            var inputArray = input as object[];
            if (inputArray != null && inputArray.Length == 1)
                input = inputArray[0];
            Type inputType = input.GetType();
            bool result = (inputType == typeof(ExcelMissing)) ||
                          (treatEmptyAsMissing && inputType == typeof(ExcelEmpty));
            if (!result && inputType == typeof(ExcelError))
                result = (ExcelError)input == ExcelError.ExcelErrorNA;
            return result;
        }

        internal static object EnumParse(Type enumType, object obj)
        {
            object result;
            string objToString = obj.ToString().Trim();
            try
            {
                result = Enum.Parse(enumType, objToString, true);
            }
            catch (ArgumentException)
            {
                throw new ArgumentException($"'{objToString}' is not a value of enum '{enumType.Name}'. Legal values are: {string.Join(", ", enumType.GetEnumNames())}");
            }
            return result;
        }

        static LambdaExpression EnumStringConversion(Type type, ExcelParameter paramReg)
        {
            // Decide whether to return a conversion function for this parameter
            if (!type.IsEnum)
                return null;

            var input = Expression.Parameter(typeof(object), "input");
            var enumTypeParam = Expression.Parameter(typeof(Type), "enumType");
            Expression<Func<Type, object, object>> enumParse = (t, s) => EnumParse(t, s);
            var result =
                Expression.Lambda(
                    Expression.Convert(
                        Expression.Invoke(enumParse, Expression.Constant(type), input),
                        type),
                    input);
            return result;
        }
    }
}
