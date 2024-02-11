using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal static class MapArrayFunctionRegistration
    {
        /// <summary>
        /// Modifies RegistrationEntries which have [ExcelMapArrayFunction],
        /// converting IEnumerable parameters to and from Excel Ranges (i.e. object[,]).
        /// This allows idiomatic .NET functions (which use sequences and lists) to be used as UDFs.
        /// 
        /// Supports the use of Excel Array formulae where a UDF returns an enumerable.
        /// 
        /// 1-dimensional Excel arrays are mapped automatically to/from IEnumerable.
        /// 2-dimensional Excel arrays can be mapped to a single function parameter with
        /// [ExcelMapPropertiesToColumnHeaders].
        /// </summary>
        public static IEnumerable<ExcelFunction> ProcessMapArrayFunctions(
            this IEnumerable<ExcelFunction> registrations,
            ParameterConversionConfiguration config = null)
        {
            foreach (var reg in registrations)
            {
                if (!(reg.FunctionAttribute is ExcelMapArrayFunctionAttribute))
                {
                    // Not considered at all
                    yield return reg;
                    continue;
                }

                try
                {
                    var inputShimParameters = reg.FunctionLambda.Parameters.ZipSameLengths(reg.ParameterRegistrations,
                                        (p, r) => new ShimParameter(p.Type, r, config)).ToList();
                    var resultShimParameter = new ShimParameter(reg.FunctionLambda.ReturnType, reg.ReturnRegistration, config);

                    // create the shim function as a lambda, using reflection
                    LambdaExpression shim = MakeObjectArrayShim(
                        reg.FunctionLambda,
                        inputShimParameters,
                        resultShimParameter);

                    // create a description of the function, with a list of the output fields
                    string functionDescription = "Returns " + resultShimParameter.HelpString;

                    // create a description of each parameter, with a list of the input fields
                    var parameterDescriptions = inputShimParameters.Select(shimParameter => "Input " +
                                                       shimParameter.HelpString).ToArray();

                    // all ok so far - modify the registration
                    reg.FunctionLambda = shim;
                    if (String.IsNullOrEmpty(reg.FunctionAttribute.Description))
                        reg.FunctionAttribute.Description = functionDescription;
                    for (int param = 0; param != reg.ParameterRegistrations.Count; ++param)
                    {
                        if (String.IsNullOrEmpty(reg.ParameterRegistrations[param].ArgumentAttribute.Description))
                            reg.ParameterRegistrations[param].ArgumentAttribute.Description =
                                parameterDescriptions[param];
                    }
                }
                catch
                {
                    // failed to shim, just pass on the original
                }
                yield return reg;
            }
        }

        private delegate object ParamsDelegate(params object[] args);

        /// <summary>
        /// Function which creates a shim for a target method.
        /// The target method is expected to take 1 or more enumerables of various types, and return a single enumerable of another type.
        /// The shim is a lambda expression which takes 1 or more object[,] parameters, and returns a single object[,]
        /// The first row of each array defines the field names, which are mapped to the public properties of the
        /// input and return types.
        /// </summary>
        /// <returns></returns>
        private static LambdaExpression MakeObjectArrayShim(
            LambdaExpression targetMethod,
            IList<ShimParameter> inputShimParameters,
            ShimParameter resultShimParameter)
        {
            var nParams = targetMethod.Parameters.Count;

            var compiledTargetMethod = targetMethod.Compile();

            // create a delegate, object*n -> object
            // (simpler, but probably slower, alternative to building it all out of Expressions)
            ParamsDelegate shimDelegate = inputObjectArray =>
            {
                try
                {
                    if (inputObjectArray.GetLength(0) != nParams)
                        throw new InvalidOperationException(
                            $"Expected {nParams} params, received {inputObjectArray.GetLength(0)}");

                    var targetMethodInputs = new object[nParams];

                    for (int i = 0; i != nParams; ++i)
                    {
                        try
                        {
                            targetMethodInputs[i] = inputShimParameters[i].ConvertShimToTarget(inputObjectArray[i]);
                        }
                        catch (Exception e)
                        {
                            throw new InvalidOperationException($"Failed to convert parameter {i + 1}: {e.Message}");
                        }
                    }

                    var targetMethodResult = compiledTargetMethod.DynamicInvoke(targetMethodInputs);

                    return resultShimParameter.ConvertTargetToShim(targetMethodResult);
                }
                catch (Exception e)
                {
                    return new object[,] { { ExcelError.ExcelErrorValue }, { e.Message } };
                }
            };

            // convert the delegate back to a LambdaExpression
            var args = targetMethod.Parameters.Select(param => Expression.Parameter(typeof(object))).ToList();
            var paramsParam = Expression.NewArrayInit(typeof(object), args);
            var closure = Expression.Constant(shimDelegate.Target);
            var call = Expression.Call(closure, shimDelegate.Method, paramsParam);
            return Expression.Lambda(call, args);
        }

        /// <summary>
        /// Class which does the work of translating a parameter or return value
        /// between the shim (e.g. object[,]) and the target (e.g. IEnumerable<typeparamref name="T"/>)
        /// </summary>
        private class ShimParameter
        {
            /// <summary>
            /// The type of the target function's parameter or return value
            /// </summary>
            private Type Type { set; get; }

            /// <summary>
            /// If the target function's parameter or return type is an IEnumerable, the enumerated type.
            /// Else null
            /// </summary>
            private Type EnumeratedType { set; get; }

            /// <summary>
            /// If the target function's parameter or return type is marked with [ExcelMapPropertiesToColumnHeaders]
            /// and mapping was successful, then this is a list of the properties to map.
            /// Else null.
            /// </summary>
            private PropertyInfo[] MappedProperties { set; get; }

            /// <summary>
            /// Construct a ShimParameter for a single parameter, or the return value, of the target function.
            /// Regular types are passed straight through with no conversion.
            /// Enumerables are converted to/from object[,] arrays.
            /// </summary>
            /// <param name="type">The type of the target function's parameter or return value</param>
            /// <param name="customAttributes">List of custom attributes defined on the
            /// target function's parameter or return value</param>
            public ShimParameter(Type type, IEnumerable<object> customAttributes)
            {
                Type = type;

                // special case - don't treat strings as enumerables
                if (type == typeof(string))
                    return;

                Type enumerable = type == typeof(IEnumerable)
                    ? type
                    : type.GetInterface(typeof(IEnumerable).Name);
                Type genericEnumerable = type.Name == typeof(IEnumerable<>).Name
                    ? type
                    : type.GetInterface(typeof(IEnumerable<>).Name);
                if (enumerable == null && genericEnumerable == null)
                    return;

                // support non-generic IEnumerables
                if (genericEnumerable == null)
                {
                    EnumeratedType = typeof(object);
                    return;
                }

                var typeArgs = genericEnumerable.GetGenericArguments();
                if (typeArgs.Length != 1)
                    return;

                EnumeratedType = typeArgs[0];

                if (!customAttributes.OfType<ExcelMapPropertiesToColumnHeadersAttribute>().Any())
                    return;

                PropertyInfo[] recordProperties =
                    EnumeratedType.GetMembers(BindingFlags.Instance | BindingFlags.Public | BindingFlags.DeclaredOnly).
                        OfType<PropertyInfo>().ToArray();
                if (recordProperties.Length == 0)
                    return;

                MappedProperties = recordProperties;
            }

            private Func<object, object>[] PropertyConverters { set; get; }

            ExcelParameter ParameterRegistration { set; get; }

            void PreparePropertyConverters<Registration>(ParameterConversionConfiguration config,
                Registration reg, Func<ParameterConversionConfiguration, Type, Registration, LambdaExpression> getConversion)
            {
                Type[] propTypes = (MappedProperties == null)
                    ? new[] { EnumeratedType ?? Type }
                    : Array.ConvertAll(MappedProperties, p => p.PropertyType);
                LambdaExpression[] lambdas = Array.ConvertAll(propTypes,
                    pt => (config == null) ? null : getConversion(config, EnumeratedType ?? Type, reg));
                lambdas = Array.ConvertAll(lambdas, l => CastParamAndResult(l, typeof(object)));
                PropertyConverters = Array.ConvertAll(lambdas,
                    l => (l == null) ? null : (Func<object, object>)l.Compile());
            }

            public ShimParameter(Type type, ExcelParameter reg, ParameterConversionConfiguration config)
                : this(type, reg.CustomAttributes)
            {
                // Try to find a converter for EnumeratedType
                ParameterRegistration = reg;
                PreparePropertyConverters(config, reg, ParameterConversionRegistration.GetParameterConversion);
            }

            ExcelReturn ReturnRegistration { set; get; }

            public ShimParameter(Type type, ExcelReturn reg, ParameterConversionConfiguration config)
                : this(type, reg.CustomAttributes)
            {
                ReturnRegistration = reg;
                PreparePropertyConverters(config, reg, ParameterConversionRegistration.GetReturnConversion);
            }

            /// <summary>
            /// Returns a help string for the parameter or return type, containing
            /// a list of the fields in the header row (if appropriate).
            /// </summary>
            public string HelpString
            {
                get
                {
                    if (MappedProperties != null)
                        return "array, with header row containing:\n" + String.Join(
                            ",", this.MappedProperties.Select(prop => prop.Name));

                    if (EnumeratedType != null)
                        return "single-row or single-column array of " + EnumeratedType.Name;

                    return "value, of type " + Type.Name;
                }
            }

            /// <summary>
            /// Converts a value from the shim (e.g. object[,]) to the target (e.g. IEnumerable)
            /// </summary>
            /// <param name="inputObject"></param>
            /// <returns></returns>
            public object ConvertShimToTarget(object inputObject)
            {
                if (MappedProperties != null)
                {
                    var objectArray2D = inputObject as object[,];
                    if (objectArray2D == null || objectArray2D.GetLength(0) == 0)
                        throw new ArgumentException("objectArray");

                    // extract nrows and ncols for each input array
                    int nInputRows = objectArray2D.GetLength(0) - 1;
                    int nInputCols = objectArray2D.GetLength(1);

                    // Decorate the input record properties with the matching
                    // column indices from the input array. We have to do this each time
                    // the shim is invoked to map column headers dynamically.
                    // Would this be better as a SelectMany?
                    var inputPropertyCols = ZipSameLengths(MappedProperties, PropertyConverters, (propInfo, converter) =>
                    {
                        int colIndex = -1;

                        for (int inputCol = 0; inputCol != nInputCols; ++inputCol)
                        {
                            var colName = objectArray2D[0, inputCol] as string;
                            if (colName == null)
                                continue;

                            if (propInfo.Name.Equals(colName, StringComparison.OrdinalIgnoreCase))
                            {
                                colIndex = inputCol;
                                break;
                            }
                        }
                        if (colIndex == -1)
                            throw new InvalidOperationException($"No column found for property {propInfo.Name}");
                        return Tuple.Create(propInfo, colIndex, converter);
                    }).ToArray();

                    // create a sequence of InputRecords
                    Array records = Array.CreateInstance(this.EnumeratedType, nInputRows);

                    // populate it
                    for (int row = 0; row != nInputRows; ++row)
                    {
                        object inputRecord;
                        try
                        {
                            // try using constructor which takes parameters in their declared order
                            inputRecord = Activator.CreateInstance(this.EnumeratedType,
                                inputPropertyCols.Select(
                                    prop =>
                                        ConvertFromExcelObject(objectArray2D[row + 1, prop.Item2],
                                            prop.Item1.PropertyType, prop.Item3)).ToArray());
                        }
                        catch (MissingMethodException)
                        {
                            // try a different way... default constructor and then set properties
                            inputRecord = Activator.CreateInstance(this.EnumeratedType);

                            // populate the record
                            foreach (var prop in inputPropertyCols)
                                prop.Item1.SetValue(inputRecord,
                                    ConvertFromExcelObject(objectArray2D[row + 1, prop.Item2],
                                        prop.Item1.PropertyType, prop.Item3), null);
                        }

                        records.SetValue(inputRecord, row);
                    }
                    return records;
                }

                if (EnumeratedType != null)
                {
                    // the target needs a 1 dimensional sequence of EnumeratedType, in either orientation
                    var objectArray = inputObject as object[,];
                    Array result;
                    if (objectArray == null)
                    {
                        // attempt to convert single item directly
                        result = Array.CreateInstance(EnumeratedType, 1);
                        result.SetValue(ConvertFromExcelObject(inputObject, EnumeratedType, PropertyConverters[0]), 0);
                    }
                    else
                    {
                        // cast each input object to the required type
                        int nRows = objectArray.GetLength(0);
                        int nCols = objectArray.GetLength(1);
                        if (nRows != 1 && nCols != 1)
                            throw new InvalidOperationException("A 1 dimensional array is required");

                        // create the required concrete array type
                        int nItems = nRows == 1 ? nCols : nRows;
                        result = Array.CreateInstance(EnumeratedType, nItems);
                        for (int i = 0; i != nItems; ++i)
                            result.SetValue(ConvertFromExcelObject(objectArray[nRows == 1 ? 0 : i, nRows == 1 ? i : 0], this.EnumeratedType, PropertyConverters[0]), i);
                    }
                    return result;
                }

                // it's a simple value type
                return ConvertFromExcelObject(inputObject, this.Type, PropertyConverters[0]);
            }

            /// <summary>
            /// Converts a value from the target (e.g. IEnumerable) to the shim (e.g. object[,])
            /// </summary>
            /// <param name="outputObject"></param>
            /// <returns></returns>
            public object ConvertTargetToShim(object outputObject)
            {
                if (MappedProperties != null)
                {
                    var genericToArray =
                        typeof(Enumerable).GetMethods(BindingFlags.Static | BindingFlags.Public)
                            .First(mi => mi.Name == "ToArray");
                    if (genericToArray == null)
                        throw new InvalidOperationException("Internal error. Failed to find Enumerable.ToArray");
                    var toArray = genericToArray.MakeGenericMethod(this.EnumeratedType);
                    var returnRecordArray = toArray.Invoke(null, new object[] { outputObject }) as Array;
                    if (returnRecordArray == null)
                        throw new InvalidOperationException("Internal error. Failed to convert return record to Array");

                    // create a return object array and populate the first row
                    var nReturnRows = returnRecordArray.Length;
                    var returnObjectArray = new object[nReturnRows + 1, this.MappedProperties.Length];
                    for (int outputCol = 0; outputCol != this.MappedProperties.Length; ++outputCol)
                        returnObjectArray[0, outputCol] = this.MappedProperties[outputCol].Name;

                    // iterate through the entire array and populate the output
                    for (int returnRow = 0; returnRow != nReturnRows; ++returnRow)
                    {
                        for (int returnCol = 0; returnCol != this.MappedProperties.Length; ++returnCol)
                        {
                            object value = this.MappedProperties[returnCol].
                                GetValue(returnRecordArray.GetValue(returnRow), null);
                            Func<object, object> converter = PropertyConverters[returnCol];
                            if (converter != null)
                                value = converter(value);
                            returnObjectArray[returnRow + 1, returnCol] = value;
                        }
                    }

                    return returnObjectArray;
                }

                if (EnumeratedType != null)
                {
                    var genericToArray2D =
                        typeof(MapArrayFunctionRegistration).GetMethods(BindingFlags.Static | BindingFlags.Public)
                            .First(mi => mi.Name == "ToArray2D");
                    if (genericToArray2D == null)
                        throw new InvalidOperationException("Internal error. Failed to find Enumerable.ToArray2D extension method");
                    var toArray2D = genericToArray2D.MakeGenericMethod(this.EnumeratedType);
                    var returnRecordArray = toArray2D.Invoke(null, new object[] { outputObject, Orientation.Vertical, PropertyConverters[0] }) as object[,];
                    if (returnRecordArray == null)
                        throw new InvalidOperationException("Internal error. Failed to convert return record to 2D Array");

                    return returnRecordArray;
                }

                return outputObject;
            }

            /// <summary>
            /// Wrapper for Convert.ChangeType which understands Excel's use of doubles as OADates.
            /// </summary>
            /// <param name="from">Excel object to convert into a different .NET type</param>
            /// <param name="toType">Type to convert to</param>
            /// <returns>Converted object</returns>
            private static object ConvertFromExcelObject(object from, Type toType, Func<object, object> converter)
            {
                if (converter != null)
                {
                    return converter(from);
                }
                // special case when converting from Excel double to DateTime
                // no need for special case in reverse, because Excel-DNA understands a DateTime object
                if (toType == typeof(DateTime) && (from is double))
                {
                    return DateTime.FromOADate((double)from);
                }
                if (toType == typeof(DateTime) && (from is int))
                {
                    return DateTime.FromOADate((int)from);
                }
                if (from is ExcelEmpty)
                {
                    // use default ctor if it exists exist
                    if (toType.IsValueType)
                        return Activator.CreateInstance(toType);
                    if (toType == typeof(string))
                        return String.Empty;
                    if (toType.GetConstructor(Type.EmptyTypes) != null)
                        return Activator.CreateInstance(toType);
                    return null;
                }
                return Convert.ChangeType(from, toType);
            }
        }

        //////////////////////////////////////////////////////////////////////////////////////////////
        #region Helper Methods

        /// <summary>
        /// Same as Zip except throws an exception if not same length
        /// </summary>
        public static IEnumerable<TResult> ZipSameLengths<TFirst, TSecond, TResult>(
            this IEnumerable<TFirst> first,
            IEnumerable<TSecond> second,
            Func<TFirst, TSecond, TResult> resultSelector)
        {
            if (first == null) throw new ArgumentNullException("first");
            if (second == null) throw new ArgumentNullException("second");
            if (resultSelector == null) throw new ArgumentNullException("resultSelector");

            using (var enum1 = first.GetEnumerator())
            using (var enum2 = second.GetEnumerator())
            {
                while (enum1.MoveNext())
                {
                    if (enum2.MoveNext())
                        yield return resultSelector(enum1.Current, enum2.Current);
                    else
                        throw new InvalidOperationException("First sequence had more elements than second");
                }
                if (enum2.MoveNext())
                    throw new InvalidOperationException("Second sequence had more elements than first");
            }
        }

        public enum Orientation
        {
            Horizontal,
            Vertical
        }

        public static object[,] ToArray2D<T>(this IEnumerable<T> input, Orientation orient, Func<object, object> itemConverter = null)
        {
            var list = input.ToList();
            var result =
                new object[orient == Orientation.Horizontal ? 1 : list.Count,
                    orient == Orientation.Horizontal ? list.Count : 1];
            if (itemConverter == null)
                itemConverter = (x) => (object)x;
            for (int i = 0; i != list.Count; ++i)
            {
                result[orient == Orientation.Horizontal ? 0 : i,
                    orient == Orientation.Horizontal ? i : 0] = itemConverter(list[i]);
            }
            return result;
        }

        private static LambdaExpression CastResult(LambdaExpression lambda, Type type)
        {
            var result = (lambda == null)
                ? null
                : lambda.ReturnType == type
                    ? lambda
                    : Expression.Lambda(
                        Expression.Convert(Expression.Invoke(lambda, lambda.Parameters), type),
                        lambda.Parameters);
            return result;
        }

        private static LambdaExpression CastParameter(LambdaExpression lambda, Type type)
        {
            LambdaExpression result = null;
            if (lambda != null)
            {
                if (lambda.Parameters[0].Type == type)
                    result = lambda;
                else
                {
                    var input = Expression.Parameter(type, "input");
                    result =
                        Expression.Lambda(
                            Expression.Invoke(lambda, Expression.Convert(input, lambda.Parameters[0].Type)), input);
                }
            }
            return result;
        }

        private static LambdaExpression CastParamAndResult(LambdaExpression lambda, Type t)
        {
            return CastResult(CastParameter(lambda, t), t);
        }

        private static Func<object, object> GetItemConverter(LambdaExpression lambda)
        {
            Func<object, object> result = null;
            if (lambda != null)
            {
                lambda = CastParamAndResult(lambda, typeof(object));
                result = (Func<object, object>)lambda.Compile();
            }
            return result;
        }

        #endregion
    }
}
