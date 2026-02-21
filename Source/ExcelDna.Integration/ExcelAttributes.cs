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

        public ExcelFunctionAttribute(ExcelFunctionAttribute src)
        {
            Category = src.Category;
            Name = src.Name;
            Description = src.Description;
            HelpTopic = src.HelpTopic;
            IsVolatile = src.IsVolatile;
            IsHidden = src.IsHidden;
            IsExceptionSafe = src.IsExceptionSafe;
            IsMacroType = src.IsMacroType;
            IsThreadSafe = src.IsThreadSafe;
            IsClusterSafe = src.IsClusterSafe;
            ExplicitRegistration = src.ExplicitRegistration;
            SuppressOverwriteError = src.SuppressOverwriteError;
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
    /// For the arguments of object handles.
    /// </summary>
    [AttributeUsage(AttributeTargets.Parameter | AttributeTargets.ReturnValue | AttributeTargets.Class | AttributeTargets.Struct, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelHandleAttribute : Attribute
    {
        public string DisplayName = null;
    }

    /// <summary>
    /// To indicate that a type will be marshalled as object handles.
    /// </summary>
    [AttributeUsage(AttributeTargets.Assembly, AllowMultiple = true)]
    [MeansImplicitUse]
    public class ExcelHandleExternalAttribute : Attribute
    {
        public ExcelHandleExternalAttribute(Type type)
        {
            Type = type;
        }

        public Type Type { get; }
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
    /// For user-defined parameter conversions.
    /// </summary>
	[AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelParameterConversionAttribute : Attribute
    {
    }

    /// <summary>
    /// For user-defined return conversions.
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelReturnConversionAttribute : Attribute
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

    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelFunctionProcessorAttribute : Attribute
    {
    }
}
