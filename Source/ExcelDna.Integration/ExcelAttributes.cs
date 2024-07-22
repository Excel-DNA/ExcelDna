//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using JetBrains.Annotations;

namespace ExcelDna.Integration
{
    /// <summary>
    /// For user-defined functions.
    /// </summary>
	[AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelFunctionAttribute : Attribute
    {
        /// <summary>
        /// By default the name of the add-in.
        /// </summary>
        public string Category = null;

        public string Name = null;
        public string Description = null;
        public string HelpTopic = null;
        public bool IsVolatile = false;
        public bool IsHidden = false;
        public bool IsExceptionSafe = false;
        public bool IsMacroType = false;
        public bool IsThreadSafe = false;
        public bool IsClusterSafe = false;
        public bool ExplicitRegistration = false;
        public bool SuppressOverwriteError = false;

        public ExcelFunctionAttribute()
        {
        }

        public ExcelFunctionAttribute(string description)
        {
            Description = description;
        }
    }

    /// <summary>
    /// For the arguments of user-defined functions.
    /// </summary>
    [AttributeUsage(AttributeTargets.Parameter, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelArgumentAttribute : Attribute
    {
        /// <summary>
        /// Arguments of type object may receive ExcelReference.
        /// </summary>
        public bool AllowReference = false;

        public string Name = null;
        public string Description = null;

        public ExcelArgumentAttribute()
        {
        }

        public ExcelArgumentAttribute(string description)
        {
            Description = description;
        }
    }

    /// <summary>
    /// For macro commands.
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelCommandAttribute : Attribute
    {
        public string Name = null;
        public string Description = null;
        public string HelpTopic = null;
        public string ShortCut = null;
        public string MenuName = null;
        public string MenuText = null;
        public bool IsExceptionSafe = false;
        public bool ExplicitRegistration = false;
        public bool SuppressOverwriteError = false;

        [Obsolete("ExcelFunctions can be declared hidden, not ExcelCommands.")]
        public bool IsHidden = false;

        public ExcelCommandAttribute()
        {
        }

        public ExcelCommandAttribute(string description)
        {
            Description = description;
        }
    }

    /// <summary>
    /// Attribute for functions that will be mapped to an Excel UDF,
    /// using property reflection to convert Excel arrays to/from .NET enumerables.
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelMapArrayFunctionAttribute : ExcelFunctionAttribute
    {
    }

    // An extension of the ExcelFunction attribute to identify functions that should be registered as async
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelAsyncFunctionAttribute : ExcelFunctionAttribute
    {
    }

    /// <summary>
    /// Optional attribute for parameters and return values of an [ExcelMapArrayFunction] function.
    /// An enumerable of records is mapped to an Excel array, where the first row of the array contains
    /// column headers which correspond to the public properties of the record type.
    ///
    /// E.g.
    ///     struct Output { int Out; }
    ///     struct Input  { int In1; int In2; }
    ///     IEnumerable MyFunc(IEnumerable) { ... }
    /// In Excel, use an Array Formula, e.g.
    ///       | A       B       C       
    ///     --+-------------------------
    ///     1 | In1     In2     {=MyFunc(A1:B3)} -> Out
    ///     2 | 1.0     2.0     {=MyFunc(A1:B3)} -> 1.5
    ///     3 | 2.0     3.0     {=MyFunc(A1:B3)} -> 2.5
    /// </summary>
    [AttributeUsage(AttributeTargets.Parameter | AttributeTargets.ReturnValue, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelMapPropertiesToColumnHeadersAttribute : Attribute
    {
    }

    /// <summary>
    /// For user-defined parameter conversions.
    /// </summary>
	[AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelParameterConversionAttribute : Attribute
    {
    }

    /// <summary>
    /// For user-defined function execution handlers.
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelFunctionExecutionHandlerSelectorAttribute : Attribute
    {
    }
}
