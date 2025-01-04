using ExcelDna.Integration;
using ExcelDna.Registration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDna.AddIn.RegistrationSample
{
    public class ExampleAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex => "!!! ERROR: " + ex.ToString());

            // Set the Parameter Conversions before they are applied by the ProcessParameterConversions call below.
            // CONSIDER: We might change the registration to be an object...?
            var conversionConfig = GetParameterConversionConfig();

            var functionHandlerConfig = GetFunctionExecutionHandlerConfig();

            ExcelRegistration.GetExcelFunctions()
                .ProcessMapArrayFunctions(conversionConfig)
                .ProcessParameterConversions(conversionConfig)
                .ProcessAsyncRegistrations(nativeAsyncIfAvailable: false)
                .ProcessParamsRegistrations()
                .ProcessFunctionExecutionHandlers(functionHandlerConfig)
                .RegisterFunctions()
                ;

            // First example if Instance -> Static conversion
            InstanceMemberRegistration.TestInstanceRegistration();
        }

        public void AutoClose()
        {
        }

        static ParameterConversionConfiguration GetParameterConversionConfig()
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

                // Register some type conversions (note the ordering discussed above)        
                .AddParameterConversion((TestType1 value) => new TestType2(value))
                .AddParameterConversion((string value) => new TestType1(value))

                // This is a conversion applied to the return value of the function
                .AddReturnConversion((TestType1 value) => value.ToString())
                .AddReturnConversion((Complex value) => new double[2] { value.Real, value.Imaginary })

                //  .AddParameterConversion((string value) => convert2(convert1(value)));

                // This parameter conversion adds support for string[] parameters (by accepting object[] instead).
                // It uses the TypeConversion utility class defined in ExcelDna.Registration to get an object->string
                // conversion that is consist with Excel (in this case, Excel is called to do the conversion).
                .AddParameterConversion((object[] inputs) => inputs.Select(TypeConversion.ConvertToString).ToArray())

                // This is a pair of very generic conversions for Enum types
                .AddReturnConversion((Enum value) => value.ToString(), handleSubTypes: true)
                .AddParameterConversion(ParameterConversions.GetEnumStringConversion())

                .AddParameterConversion((object[] input) => new Complex(TypeConversion.ConvertToDouble(input[0]), TypeConversion.ConvertToDouble(input[1])))
                .AddNullableConversion(treatEmptyAsMissing: true, treatNAErrorAsMissing: true);

            return paramConversionConfig;
        }

        static FunctionExecutionConfiguration GetFunctionExecutionHandlerConfig()
        {
            return new FunctionExecutionConfiguration()
                .AddFunctionExecutionHandler(FunctionLoggingHandler.LoggingHandlerSelector)
                .AddFunctionExecutionHandler(CacheFunctionExecutionHandler.CacheHandlerSelector)
                .AddFunctionExecutionHandler(TimingFunctionExecutionHandler.TimingHandlerSelector)
                .AddFunctionExecutionHandler(SuppressInDialogFunctionExecutionHandler.SuppressInDialogSelector);
        }
    }
}
