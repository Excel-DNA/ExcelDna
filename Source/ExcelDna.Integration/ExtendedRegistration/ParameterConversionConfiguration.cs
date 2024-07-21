using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal class ParameterConversionConfiguration
    {
        internal class ParameterConversion
        {
            // Conversion receives the parameter type and parameter registration info, 
            // and should return an Expression<Func<TTo, TFrom>> 
            // (and may optionally update the information in the ExcelParameterRegistration.
            // May return null to indicate that no conversion should be applied.
            public Func<Type, IExcelFunctionParameter, LambdaExpression> Conversion { get; private set; }

            // The TypeFilter is used as a quick filter to decide whether the Conversion function should be called for a parameter.
            // TypeFilter may be null to indicate that conversion should be applied for all types.
            // (The Conversion function may anyway return null to indicate that no conversion should be applied.)
            public Type TypeFilter { get; private set; }

            public ParameterConversion(Func<Type, IExcelFunctionParameter, LambdaExpression> conversion, Type typeFilter = null)
            {
                if (conversion == null)
                    throw new ArgumentNullException("conversion");

                Conversion = conversion;
                TypeFilter = typeFilter;
            }

            internal LambdaExpression Convert(Type paramType, IExcelFunctionParameter paramReg)
            {
                if (TypeFilter != null && paramType != TypeFilter)
                    return null;

                return Conversion(paramType, paramReg);
            }
        }

        internal class ReturnConversion
        {
            // Conversion receives the return type and list of custom attributes applied to the return value,
            // and should return  an Expression<Func<TTo, TFrom>> 
            // (and may optionally update the information in the ExcelParameterRegistration.
            // May return null to indicate that no conversion should be applied.
            public Func<Type, IExcelFunctionReturn, LambdaExpression> Conversion { get; private set; }

            // TypeFilter is used as a quick filter to decide whether the conversion function should be called for a return value.
            // TypeFilter be null to indicate that conversion should be applied for all types
            // The Conversion function may anyway return null to indicate that no conversion should be applied.
            public Type TypeFilter { get; private set; }

            /// <summary>
            /// If true, the conversion will also convert all subtypes of its input type
            /// </summary>
            public bool HandleSubTypes { get; private set; }

            public ReturnConversion(Func<Type, IExcelFunctionReturn, LambdaExpression> conversion, Type typeFilter = null, bool handleSubTypes = false)
            {
                if (conversion == null)
                    throw new ArgumentNullException("conversion");

                Conversion = conversion;
                TypeFilter = typeFilter;
                HandleSubTypes = handleSubTypes;
            }

            internal LambdaExpression Convert(Type returnType, IExcelFunctionReturn returnRegistration)
            {
                if (TypeFilter != null && returnType != TypeFilter && (!HandleSubTypes || !returnType.IsSubclassOf(TypeFilter)))
                    return null;

                LambdaExpression result = Conversion(returnType, returnRegistration);

                if (TypeFilter != null && returnType != TypeFilter)
                {
                    var returnValue = Expression.Parameter(returnType, "returnValue");
                    var castExpr = Expression.Convert(returnValue, TypeFilter);
                    var composeExpr = Expression.Invoke(result, castExpr);
                    result = Expression.Lambda(composeExpr, returnValue);
                }
                return result;
            }
        }

        internal List<ParameterConversion> ParameterConversions { get; private set; }
        internal List<ReturnConversion> ReturnConversions { get; private set; }

        public ParameterConversionConfiguration()
        {
            ParameterConversions = new List<ParameterConversion>();
            ReturnConversions = new List<ReturnConversion>();
        }

        #region Various overloads for adding conversions

        // Most general case - called by the overloads below
        /// <summary>
        /// Converts a parameter from an Excel-friendly type (e.g. object, or string) to an add-in friendly type, e.g. double? or InternalType.
        /// Will only be considered for those parameters that have a 'to' type that matches targetTypeOrNull,
        ///  or for all types if null is passes for the first parameter.
        /// </summary>
        /// <param name="parameterConversion"></param>
        /// <param name="targetTypeOrNull"></param>
        public ParameterConversionConfiguration AddParameterConversion(Func<Type, IExcelFunctionParameter, LambdaExpression> parameterConversion, Type targetTypeOrNull = null)
        {
            var pc = new ParameterConversion(parameterConversion, targetTypeOrNull);
            ParameterConversions.Add(pc);
            return this;
        }

        public ParameterConversionConfiguration AddParameterConversion<TTo>(Func<Type, IExcelFunctionParameter, LambdaExpression> parameterConversion)
        {
            AddParameterConversion(parameterConversion, typeof(TTo));
            return this;
        }

        public ParameterConversionConfiguration AddParameterConversion<TFrom, TTo>(Expression<Func<TFrom, TTo>> convert)
        {
            AddParameterConversion<TTo>((unusedParamType, unusedParamReg) => convert);
            return this;
        }

        public ParameterConversionConfiguration AddParameterConversions(IEnumerable<Func<Type, IExcelFunctionParameter, LambdaExpression>> parameterConversions)
        {
            foreach (var i in parameterConversions)
                AddParameterConversion(i);

            return this;
        }

        // Most general case - called by the overloads below
        public ParameterConversionConfiguration AddReturnConversion(Func<Type, IExcelFunctionReturn, LambdaExpression> returnConversion, Type targetTypeOrNull = null, bool handleSubTypes = false)
        {
            var rc = new ReturnConversion(returnConversion, targetTypeOrNull, handleSubTypes);
            ReturnConversions.Add(rc);
            return this;
        }

        public ParameterConversionConfiguration AddReturnConversion<TFrom>(Func<Type, IExcelFunctionReturn, LambdaExpression> returnConversion, Type targetTypeOrNull = null, bool handleSubTypes = false)
        {
            AddReturnConversion(returnConversion, typeof(TFrom), handleSubTypes);
            return this;
        }

        public ParameterConversionConfiguration AddReturnConversion<TFrom, TTo>(Expression<Func<TFrom, TTo>> convert, bool handleSubTypes = false)
        {
            AddReturnConversion<TFrom>((unusedReturnType, unusedAttributes) => convert, null, handleSubTypes);
            return this;
        }
        #endregion

        Func<Type, IExcelFunctionParameter, LambdaExpression> GetNullableConversion(bool treatEmptyAsMissing, bool treatNAErrorAsMissing)
        {
            return (type, paramReg) => ExtendedRegistration.ParameterConversions.NullableConversion(this, type, paramReg, treatEmptyAsMissing, treatNAErrorAsMissing);
        }

        /// <summary>
        /// Adds a Nullable conversion that will also translate any type parameter T of Nullable[T] for which there is a conversion in the configutation.
        /// Note that the added rule is quite generic and only has access to the T conversion rules that have already been added before it, so you should
        /// call this at the very bottom of your configuration setup sequence.
        /// </summary>
        /// <param name="treatEmptyAsMissing">If true, any empty cells will be treated as null values</param>
        /// <param name="treatNAErrorAsMissing">If true, any #NA! errors will be treated as null values</param>
        /// <returns>The parameter conversion configuration with the new added rule</returns>
        public ParameterConversionConfiguration AddNullableConversion(bool treatEmptyAsMissing = false, bool treatNAErrorAsMissing = false)
        {
            return AddParameterConversion(GetNullableConversion(treatEmptyAsMissing, treatNAErrorAsMissing));
        }
    }
}
