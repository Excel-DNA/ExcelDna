using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using Expr = System.Linq.Expressions.Expression;

namespace ExcelDna.Integration.ExtendedRegistration
{
    // CONSIDER: Can one use an ExpressionVisitor to do these things....?
    internal static class ParameterConversionRegistration
    {
        public static IEnumerable<ExcelFunction> ProcessParameterConversions(this IEnumerable<ExcelFunction> registrations, ParameterConversionConfiguration conversionConfig)
        {
            foreach (var reg in registrations)
            {
                // Keep a list of conversions for each parameter
                // TODO: Prevent having a cycle, but allow arbitrary ordering...?

                var paramsConversions = new List<List<LambdaExpression>>();
                for (int i = 0; i < reg.FunctionLambda.Parameters.Count; i++)
                {
                    var initialParamType = reg.FunctionLambda.Parameters[i].Type;
                    var paramReg = reg.ParameterRegistrations[i];

                    var paramConversions = GetParameterConversions(conversionConfig, initialParamType, paramReg);
                    paramsConversions.Add(paramConversions);
                } // for each parameter !

                // Process return conversions
                var returnConversions = GetReturnConversions(conversionConfig, reg.FunctionLambda.ReturnType, reg.ReturnRegistration);

                // Now we apply all the conversions
                ApplyConversions(reg, paramsConversions, returnConversions);

                yield return reg;
            }
        }

        // returnsConversion and the entries in paramsConversions may be null.
        public static void ApplyConversions(ExcelFunction reg, List<List<LambdaExpression>> paramsConversions, List<LambdaExpression> returnConversions)
        {
            // CAREFUL: The parameter transformations are applied in reverse order to how they're identified.
            // We do the following transformation
            //      public static string dnaParameterConvertTest(double? optTest) {   };
            //
            // with conversions convert1 and convert2 taking us from Type1 to double?
            // 
            // to
            //      public static string dnaParameterConvertTest(Type1 optTest) 
            //      {   
            //          return convertRet2(convertRet1(
            //                      dnaParameterConvertTest(
            //                          paramConvert1(optTest)
            //                            )));
            //      };
            // 
            // and then with a conversion from object to Type1, resulting in
            //
            //      public static string dnaParameterConvertTest(object optTest) 
            //      {   
            //          return convertRet2(convertRet1(
            //                      dnaParameterConvertTest(
            //                          paramConvert1(paramConvert2(optTest))
            //                            )));
            //      };

            // build up the invoke expression for each parameter
            var wrappingParameters = reg.FunctionLambda.Parameters.Select(p => Expression.Parameter(p.Type, p.Name)).ToList();

            // Build the nested parameter convertion expression.
            // Start with the wrapping parameters as they are. Then replace with the nesting of conversions as needed.
            var paramExprs = new List<Expression>(wrappingParameters);

            if (paramsConversions != null)
            {
                Debug.Assert(reg.FunctionLambda.Parameters.Count == paramsConversions.Count);

                // NOTE: To cater for the Range COM type equivalance, we need to distinguish the FunctionLambda's parameter type and the paramConversion ReturnType.
                //       These need not be the same, but the should at least be equivalent.

                for (int i = 0; i < paramsConversions.Count; i++)
                {
                    var paramConversions = paramsConversions[i];
                    if (paramConversions == null)
                        continue;

                    // If we have a list, there should be at least one conversion in it.
                    Debug.Assert(paramConversions.Count > 0);
                    // Update the calling parameter type to be the outer one in the conversion chain.
                    wrappingParameters[i] = Expr.Parameter(paramConversions.Last().Parameters[0].Type, wrappingParameters[i].Name);
                    // Start with just the (now updated) outer param which will be the inner-most value in the conversion chain
                    Expression wrappedExpr = wrappingParameters[i];
                    // Need to go in reverse for the parameter wrapping
                    // Need to now build from the inside out
                    foreach (var conversion in Enumerable.Reverse(paramConversions))
                    {
                        wrappedExpr = Expr.Invoke(conversion, wrappedExpr);
                    }
                    paramExprs[i] = wrappedExpr;
                }
            }

            var wrappingCall = Expr.Invoke(reg.FunctionLambda, paramExprs);
            if (returnConversions != null)
            {
                foreach (var conversion in returnConversions)
                    wrappingCall = Expr.Invoke(conversion, wrappingCall);
            }

            reg.FunctionLambda = Expr.Lambda(wrappingCall, reg.FunctionLambda.Name, wrappingParameters);
        }

        static LambdaExpression ComposeLambdas(IEnumerable<LambdaExpression> lambdas)
        {
            LambdaExpression result = null;
            if (lambdas != null)
            {
                var convsIter = lambdas.GetEnumerator();
                if (convsIter.MoveNext())
                {
                    result = convsIter.Current;
                    while (convsIter.MoveNext())
                    {
                        result = Expression.Lambda(Expression.Invoke(result, convsIter.Current),
                            convsIter.Current.Parameters);
                    }
                }
            }
            return result;
        }

        internal static LambdaExpression GetParameterConversion(ParameterConversionConfiguration conversionConfig,
            Type initialParamType, ExcelParameter paramRegistration)
        {
            return ComposeLambdas(GetParameterConversions(conversionConfig, initialParamType, paramRegistration));
        }

        // Should return null if there are no conversions to apply
        internal static List<LambdaExpression> GetParameterConversions(ParameterConversionConfiguration conversionConfig, Type initialParamType, ExcelParameter paramRegistration)
        {
            var appliedConversions = new List<LambdaExpression>();

            // paramReg might be modified internally by the conversions, but won't become a different object
            var paramType = initialParamType; // Might become a different type as we convert
            foreach (var paramConversion in conversionConfig.ParameterConversions)
            {
                var lambda = paramConversion.Convert(paramType, paramRegistration);
                if (lambda == null)
                    continue;

                // We got one to apply...
                // Some sanity checks
                Debug.Assert(lambda.Parameters.Count == 1);
                Debug.Assert(lambda.ReturnType == paramType || lambda.ReturnType.IsEquivalentTo(paramType));

                appliedConversions.Add(lambda);

                // Change the Parameter Type to be whatever the conversion function takes us to
                // for the next round of processing
                paramType = lambda.Parameters[0].Type;
            }

            if (appliedConversions.Count == 0)
                return null;

            return appliedConversions;
        }

        internal static LambdaExpression GetReturnConversion(ParameterConversionConfiguration conversionConfig,
            Type initialReturnType, ExcelReturn returnRegistration)
        {
            return ComposeLambdas(GetReturnConversions(conversionConfig, initialReturnType, returnRegistration));
        }

        internal static List<LambdaExpression> GetReturnConversions(ParameterConversionConfiguration conversionConfig, Type initialReturnType, ExcelReturn returnRegistration)
        {
            return GetReturnConversions(conversionConfig.ReturnConversions, initialReturnType, returnRegistration);
        }

        internal static List<LambdaExpression> GetReturnConversions(List<ParameterConversionConfiguration.ReturnConversion> returnConversions, Type initialReturnType, ExcelReturn returnRegistration)
        {
            var appliedConversions = new List<LambdaExpression>();

            // paramReg might be modified internally by the conversions, but won't become a different object
            var returnType = initialReturnType; // Might become a different type as we convert

            foreach (var returnConversion in returnConversions)
            {
                var lambda = returnConversion.Convert(returnType, returnRegistration);
                if (lambda == null)
                    continue;

                // We got one to apply...
                // Some sanity checks
                Debug.Assert(lambda.Parameters.Count == 1);
                Debug.Assert(lambda.Parameters[0].Type == returnType);

                appliedConversions.Add(lambda);

                // Change the Return Type to be whatever the conversion function returns
                // for the next round of processing
                returnType = lambda.ReturnType;
            }

            if (appliedConversions.Count == 0)
                return null;

            return appliedConversions;
        }
    }
}
