using System;
using System.Linq;
using ExcelDna.Integration;

namespace ExcelDna.AddIn.RegistrationSample
{
    public static class ParamsFunctionExamples
    {

        // This function has its final argument marked with 'params' 
        // Via the Registration helper will be registered in Excel as a function with 29 or 125 arguments,
        // and the wrapper will automatically remove 'ExcelMissing' values.
        //
        // If ExplicitRegistration="true" was _not_ in the .dna file, then
        // this function would normally be registed automatically by Excel-DNA.
        // (without the params processing) before being registered again here with the params expansion.
        //
        // We would prevent that by adding the ExplicitRegistration=true flag.
        // (But in this example it's redundant, since the .dna file already protects this.)
        // 
        // Check how the parameters and their descriptions appear in the Function Arguments dialog...
        [ExcelFunction(ExplicitRegistration = true)]
        public static string dnaParamsFunc(
            [ExcelArgument(Name = "first.Input", Description = "is a useful start")]
            object input,
            [ExcelArgument(Description = "is another param start")]
            string QtherInpEt,
            [ExcelArgument(Name = "Value", Description = "gives the Rest")]
            params object[] args)
        {
            return input + "," + QtherInpEt + ", : " + args.Length;
        }

        // When we enter =dnaParamsFunc2("a",,"c","d",,"f") we expect a,,c,[2:d,ExcelMissing,f]

        [ExcelFunction(ExplicitRegistration = true)]
        public static string dnaParamsFunc2(
            [ExcelArgument(Name = "first.Input", Description = "is a useful start")]
            object input,
            [ExcelArgument(Name = "second.Input", Description = "is some more stuff")]
            string input2,
            [ExcelArgument(Description = "is another param ")]
            string QtherInpEt,
            [ExcelArgument(Name = "Value", Description = "gives the Rest")]
            params object[] args)
        {
            return input + "," + input2 + "," + QtherInpEt + ", " + PrintArray(args);
        }


        [ExcelFunction]
        public static string dnaJoinStringParams(string separator, params string[] values)
        {
            return String.Join(separator, values);
        }

        static string PrintArray(object[] array)
        {
            var content = string.Join(",", array.Select(ValueType => ValueType.ToString()));
            return $"[{array.Length}: {content}]";
        }
    }
}
