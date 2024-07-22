using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal static class ParamsRegistration
    {
        public static bool IsParamsMethod(ExcelFunction reg)
        {
            var lastParam = reg.ParameterRegistrations.LastOrDefault();
            return lastParam != null && lastParam.CustomAttributes.Any(att => att is ParamArrayAttribute)
                && reg.FunctionLambda.Parameters.Last().Type.IsArray;
        }

        /// <summary>
        /// Adds parameters to amy function ending in a params parameter, to take the total number of parameters to 125 (or 29 under Excel 2003).
        /// </summary>
        /// <param name="registrations"></param>
        /// <returns></returns>
        public static IEnumerable<ExcelFunction> ProcessParamsRegistrations(this IEnumerable<ExcelFunction> registrations)
        {
            foreach (var reg in registrations)
            {
                try
                {
                    if (IsParamsMethod(reg))
                    {
                        reg.FunctionLambda = WrapMethodParams(reg.FunctionLambda);

                        // Clean out ParamArray attribute for the last parameter (there will be one)
                        var lastParam = reg.ParameterRegistrations.Last();
                        lastParam.CustomAttributes.RemoveAll(att => att is ParamArrayAttribute);

                        // Add more attributes for the 'params' arguments
                        // Adjust the first one from myInput to myInput1
                        var paramsArgAttrib = lastParam.ArgumentAttribute;
                        paramsArgAttrib.Name = paramsArgAttrib.Name + "1";

                        // Add the ellipse argument
                        reg.ParameterRegistrations.Add(
                            new ExcelParameter(
                                new ExcelArgumentAttribute
                                {
                                    Name = "...",
                                    Description = paramsArgAttrib.Description,
                                    AllowReference = paramsArgAttrib.AllowReference
                                }));

                        // And the rest with no Name, but copying the description
                        var restCount = reg.FunctionLambda.Parameters.Count - reg.ParameterRegistrations.Count;
                        for (int i = 0; i < restCount; i++)
                        {
                            ExcelParameter newParameter = new ExcelParameter(
                                    new ExcelArgumentAttribute
                                    {
                                        Name = string.Empty,
                                        Description = paramsArgAttrib.Description,
                                        AllowReference = paramsArgAttrib.AllowReference
                                    });
                            newParameter.CustomAttributes.AddRange(lastParam.CustomAttributes);
                            reg.ParameterRegistrations.Add(newParameter);
                        }

                        // Check that we still have a valid registration structure
                        // TODO: Make this safer...
                        Debug.Assert(reg.IsValid());
                    }
                }
                catch (Exception ex)
                {
                    Logging.LogDisplay.WriteLine("Exception while registering method {0} - {1}", reg.FunctionAttribute.Name, ex.ToString());
                    continue;
                }

                yield return reg;
            }

            //    // TODO: Check Argument length matching...

            //    var att = method.GetCustomAttribute<ExcelFunctionAsyncAttribute>();
            //    // TODO: What if it doesn't have one...?
            //    if (string.IsNullOrEmpty(att.Name)) att.Name = method.Name;

            //    var argAtts = method.GetParameters()
            //                  .Where(p => p.ParameterType != typeof(CancellationToken))
            //                  .Select(param =>
            //                  {
            //                      var argAtt = param.GetCustomAttribute<ExcelArgumentAttribute>()
            //                                   ?? new ExcelArgumentAttribute();
            //                      if (string.IsNullOrEmpty(argAtt.Name)) argAtt.Name = param.Name;
            //                      return argAtt;
            //                  })
            //                    .Cast<object>().ToList();

            //    delList.Add(del);
            //    attList.Add(att);
            //    argAttList.Add(argAtts);
            //}
            //ExcelIntegration.RegisterDelegates(delList, attList, argAttList);
        }

        public delegate TResult CustomFunc29<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16, T17, T18, T19, T20, T21, T22, T23, T24, T25, T26, T27, T28, T29, TResult>
                                     (T1 arg1, T2 arg2, T3 arg3, T4 arg4, T5 arg5, T6 arg6, T7 arg7, T8 arg8, T9 arg9, T10 arg10, T11 arg11, T12 arg12, T13 arg13, T14 arg14, T15 arg15, T16 arg16, T17 arg17, T18 arg18, T19 arg19, T20 arg20, T21 arg21, T22 arg22, T23 arg23, T24 arg24, T25 arg25, T26 arg26, T27 arg27, T28 arg28, T29 arg29);
        public delegate TResult CustomFunc125<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16, T17, T18, T19, T20, T21, T22, T23, T24, T25, T26, T27, T28, T29, T30, T31, T32, T33, T34, T35, T36, T37, T38, T39, T40, T41, T42, T43, T44, T45, T46, T47, T48, T49, T50, T51, T52, T53, T54, T55, T56, T57, T58, T59, T60, T61, T62, T63, T64, T65, T66, T67, T68, T69, T70, T71, T72, T73, T74, T75, T76, T77, T78, T79, T80, T81, T82, T83, T84, T85, T86, T87, T88, T89, T90, T91, T92, T93, T94, T95, T96, T97, T98, T99, T100, T101, T102, T103, T104, T105, T106, T107, T108, T109, T110, T111, T112, T113, T114, T115, T116, T117, T118, T119, T120, T121, T122, T123, T124, T125, TResult>
                                     (T1 arg1, T2 arg2, T3 arg3, T4 arg4, T5 arg5, T6 arg6, T7 arg7, T8 arg8, T9 arg9, T10 arg10, T11 arg11, T12 arg12, T13 arg13, T14 arg14, T15 arg15, T16 arg16, T17 arg17, T18 arg18, T19 arg19, T20 arg20, T21 arg21, T22 arg22, T23 arg23, T24 arg24, T25 arg25, T26 arg26, T27 arg27, T28 arg28, T29 arg29, T30 arg30, T31 arg31, T32 arg32, T33 arg33, T34 arg34, T35 arg35, T36 arg36, T37 arg37, T38 arg38, T39 arg39, T40 arg40, T41 arg41, T42 arg42, T43 arg43, T44 arg44, T45 arg45, T46 arg46, T47 arg47, T48 arg48, T49 arg49, T50 arg50, T51 arg51, T52 arg52, T53 arg53, T54 arg54, T55 arg55, T56 arg56, T57 arg57, T58 arg58, T59 arg59, T60 arg60, T61 arg61, T62 arg62, T63 arg63, T64 arg64, T65 arg65, T66 arg66, T67 arg67, T68 arg68, T69 arg69, T70 arg70, T71 arg71, T72 arg72, T73 arg73, T74 arg74, T75 arg75, T76 arg76, T77 arg77, T78 arg78, T79 arg79, T80 arg80, T81 arg81, T82 arg82, T83 arg83, T84 arg84, T85 arg85, T86 arg86, T87 arg87, T88 arg88, T89 arg89, T90 arg90, T91 arg91, T92 arg92, T93 arg93, T94 arg94, T95 arg95, T96 arg96, T97 arg97, T98 arg98, T99 arg99, T100 arg100, T101 arg101, T102 arg102, T103 arg103, T104 arg104, T105 arg105, T106 arg106, T107 arg107, T108 arg108, T109 arg109, T110 arg110, T111 arg111, T112 arg112, T113 arg113, T114 arg114, T115 arg115, T116 arg116, T117 arg117, T118 arg118, T119 arg119, T120 arg120, T121 arg121, T122 arg122, T123 arg123, T124 arg124, T125 arg125);

        static LambdaExpression WrapMethodParams(LambdaExpression functionLambda)
        {
            /* We are converting:
             *     [ExcelFunction(...)]
             *     public static string myFunc(string input, int otherInput, params object[] args)
             *     {    
             *          ...
             *     }
             * 
             * into:
             *     [ExcelFunction(...)]
             *     public static string myFunc(string input, int otherInput, object arg3, object arg4, object arg5, object arg6, {...until...}, object arg125)
             *     {
             *         // First we figure where in the list to stop building the param array
             *         int lastArgToAdd = 0;
             *         if (!(arg3 is ExcelMissing)) lastArgToAdd = 3;
             *         if (!(arg4 is ExcelMissing)) lastArgToAdd = 4;
             *         ...
             *         if (!(arg125 is ExcelMissing)) lastArgToAdd = 125;
             *     
             *         // Then add until we get there
             *         List<object> args = new List<object>();
             *         if (lastArgToAdd >= 3) args.Add(arg3);
             *         if (lastArgToAdd >= 4) args.Add(arg4);
             *         ...
             *         if (lastArgToAdd >= 125) args.Add(arg125);
             *        
             *         Array<object> argsArray = args.ToArray();
             *         return myFunc(input, otherInput, argsArray);
             *     }
             * 
             * 
             */

            int maxArguments;
            if (ExcelDnaUtil.ExcelVersion >= 12.0)
            {
                maxArguments = 125; // Constrained by 255 char registration string, take off 3 type chars, use up to 2 chars per param (before we start doing object...) (& also return)
                                    // CONSIDER: Might improve this if we generate the delegate based on the max length...
            }
            else
            {
                maxArguments = 29; // Or maybe 30?
            }

            var normalParams = functionLambda.Parameters.Take(functionLambda.Parameters.Count() - 1).ToList();
            var normalParamCount = normalParams.Count;
            var paramsParamCount = maxArguments - normalParamCount;
            var allParamExprs = new List<ParameterExpression>(normalParams);
            var blockExprs = new List<Expression>();
            var blockVars = new List<ParameterExpression>();

            // Run through the arguments looking for the position of the last non-ExcelMissing argument
            var lastArgVarExpr = Expression.Variable(typeof(int));
            blockVars.Add(lastArgVarExpr);
            blockExprs.Add(Expression.Assign(lastArgVarExpr, Expression.Constant(0)));
            for (int i = normalParamCount + 1; i <= maxArguments; i++)
            {
                allParamExprs.Add(Expression.Parameter(typeof(object), "arg" + i));

                var lenTestParam = Expression.IfThen(Expression.Not(Expression.TypeIs(allParamExprs[i - 1], typeof(ExcelMissing))),
                                    Expression.Assign(lastArgVarExpr, Expression.Constant(i)));
                blockExprs.Add(lenTestParam);
            }

            // We know that last parameter is an array type
            // Create a new list to hold the values
            var argsArrayType = functionLambda.Parameters.Last().Type;
            var argsType = argsArrayType.GetElementType();
            var argsListType = typeof(List<>).MakeGenericType(argsType);
            var argsListVarExpr = Expression.Variable(argsListType);
            blockVars.Add(argsListVarExpr);
            var argListAssignExpr = Expression.Assign(argsListVarExpr, Expression.New(argsListType));
            blockExprs.Add(argListAssignExpr);
            // And put the (converted) arguments into the list
            for (int i = normalParamCount + 1; i <= maxArguments; i++)
            {
                var testParam = Expression.IfThen(Expression.GreaterThanOrEqual(lastArgVarExpr, Expression.Constant(i)),
                                    Expression.Call(argsListVarExpr, "Add", null,
                                        TypeConversion.GetConversion(allParamExprs[i - 1], argsType)));

                blockExprs.Add(testParam);
            }
            var argArrayVarExpr = Expression.Variable(argsArrayType);
            blockVars.Add(argArrayVarExpr);

            var argArrayAssignExpr = Expression.Assign(argArrayVarExpr, Expression.Call(argsListVarExpr, "ToArray", null));
            blockExprs.Add(argArrayAssignExpr);

            var innerParams = new List<Expression>(normalParams) { argArrayVarExpr };
            var callInner = Expression.Invoke(functionLambda, innerParams);
            blockExprs.Add(callInner);

            var blockExpr = Expression.Block(blockVars, blockExprs);

            // Build the delegate type to return
            var allParamTypes = normalParams.Select(pi => pi.Type).ToList();
            var toAdd = maxArguments - allParamTypes.Count;
            for (int i = 0; i < toAdd; i++)
            {
                allParamTypes.Add(typeof(object));
            }
            allParamTypes.Add(functionLambda.ReturnType);

            Type delegateType;
            if (maxArguments == 125)
            {
                delegateType = typeof(CustomFunc125<,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,>)
                                    .MakeGenericType(allParamTypes.ToArray());
            }
            else // if (maxArguments == 29)
            {
                delegateType = typeof(CustomFunc29<,,,,,,,,,,,,,,,,,,,,,,,,,,,,,>)
                                    .MakeGenericType(allParamTypes.ToArray());
            }
            return Expression.Lambda(delegateType, blockExpr, allParamExprs);
        }
    }
}
