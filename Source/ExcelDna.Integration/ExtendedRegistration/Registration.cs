using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Numerics;
using ExcelDna.Integration.ObjectHandles;
using ExcelDna.Registration;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal class Registration
    {
        public class Configuration
        {
            public IEnumerable<ExcelParameterConversion> ParameterConversions { get; set; }
            public IEnumerable<ExcelReturnConversion> ReturnConversions { get; set; }
            public IEnumerable<ExcelFunctionProcessor> ExcelFunctionProcessors { get; set; }
            public IEnumerable<FunctionExecutionHandlerSelector> ExcelFunctionExecutionHandlerSelectors { get; set; }
        }

        public static void Register(IEnumerable<ExcelFunctionRegistration> functions, Configuration configuration)
        {
            Register(Process(functions, configuration));
        }

        public static void Register(IEnumerable<ExcelFunctionRegistration> functions)
        {
            functions = functions.ToList();
            var lambdas = functions.Select(reg => reg.FunctionLambda).ToList();
            var attribs = functions.Select(reg => reg.FunctionAttribute).ToList<object>();
            var argAttribs = functions.Select(reg => reg.ParameterRegistrations.Select(pr => pr.ArgumentAttribute).ToList<object>()).ToList();
            ExcelIntegration.RegisterLambdaExpressions(lambdas, attribs, argAttribs);
        }

        public static IEnumerable<ExcelFunctionRegistration> Process(IEnumerable<ExcelFunctionRegistration> functions, Configuration configuration)
        {
            // Set the Parameter Conversions before they are applied by the ProcessParameterConversions call below.
            // CONSIDER: We might change the registration to be an object...?
            var conversionConfig = GetParameterConversionConfig(configuration.ParameterConversions, configuration.ReturnConversions);

            var functionHandlerConfig = GetFunctionExecutionHandlerConfig(configuration.ExcelFunctionExecutionHandlerSelectors);

            return functions
                .UpdateRegistrationsForRangeParameters()
                .ProcessFunctionProcessors(configuration.ExcelFunctionProcessors, conversionConfig)
                .ProcessParameterConversions(conversionConfig)
                .ProcessAsyncRegistrations(nativeAsyncIfAvailable: false)
                .ProcessParamsRegistrations()
                .ProcessObjectHandles()
                .ProcessFunctionExecutionHandlers(functionHandlerConfig)
                ;
        }

        private static ParameterConversionConfiguration GetParameterConversionConfig(IEnumerable<ExcelParameterConversion> parameterConversions, IEnumerable<ExcelReturnConversion> returnConversions)
        {
            // NOTE: The parameter conversion list is processed once per parameter.
            //       Parameter conversions will apply from most inside, to most outside.
            //       So to apply a conversion chain like
            //           string -> Type1 -> Type2
            //       we need to register in the (reverse) order
            //           Type1 -> Type2
            //           string -> Type1
            //
            //       (If the registration were in the order
            //           string -> Type1
            //           Type1 -> Type2
            //       the parameter (starting as Type2) would not match the first conversion,
            //       then the second conversion (Type1 -> Type2) would be applied, and no more,
            //       leaving the parameter having Type1 (and probably not eligible for Excel registration.)
            //      
            //
            //       Return conversions are also applied from most inside to most outside.
            //       So to apply a return conversion chain like
            //           Type1 -> Type2 -> string
            //       we need to register the ReturnConversions as
            //           Type1 -> Type2 
            //           Type2 -> string
            //       

            var paramConversionConfig = new ParameterConversionConfiguration()

                // Register the Standard Parameter Conversions (with the optional switch on how to treat references to empty cells)
                .AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing: true))
                .AddParameterConversion(RangeConversion.GetRangeParameterConversion, null)

                .AddParameterConversions(ParameterConversions.GetUserParameterConversions(parameterConversions))
                .AddReturnConversions(ParameterConversions.GetUserReturnConversions(returnConversions))

                // This is a conversion applied to the return value of the function
                .AddReturnConversion((Complex value) => TypeConversion.ConvertComplexToDoubles(value))

                // This parameter conversion adds support for string[] parameters (by accepting object[] instead).
                // It uses the TypeConversion utility class to get an object->string
                // conversion that is consist with Excel (in this case, Excel is called to do the conversion).
                .AddParameterConversion((object[] inputs) => inputs.Select(TypeConversion.ConvertToString).ToArray())

                .AddParameterConversion((object[,] inputs) => TypeConversion.ConvertToString2D(inputs))

                // This is a pair of very generic conversions for Enum types
                .AddReturnConversion((Enum value) => value.ToString(), handleSubTypes: true)
                .AddParameterConversion(ParameterConversions.GetEnumStringConversion())

                .AddParameterConversion((object[] input) => new Complex(TypeConversion.ConvertToDouble(input[0]), TypeConversion.ConvertToDouble(input[1])))
                .AddNullableConversion(treatEmptyAsMissing: true, treatNAErrorAsMissing: true);

            return paramConversionConfig;
        }

        private static FunctionExecutionConfiguration GetFunctionExecutionHandlerConfig(IEnumerable<FunctionExecutionHandlerSelector> excelFunctionExecutionHandlerSelectors)
        {
            FunctionExecutionConfiguration result = new FunctionExecutionConfiguration();

            foreach (var s in excelFunctionExecutionHandlerSelectors)
            {
                result = result.AddFunctionExecutionHandler((ExcelDna.Registration.ExcelFunctionRegistration functionRegistration) => s(functionRegistration));
            }

            return result;
        }
    }
}
