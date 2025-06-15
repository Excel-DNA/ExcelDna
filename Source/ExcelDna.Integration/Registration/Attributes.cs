﻿using System;
using ExcelDna.Integration;
using JetBrains.Annotations;

namespace ExcelDna.Registration
{
    // An extension of the ExcelFunction attribute to identify functions that should be registered as async
    // By default functions are set as ExplicitRegistration=true, so marked functions will not be automatically registered
    // (this is important for 'regular' functions that should be wrapped in a Task.
    // CONSIDER: Maybe add caching options?
    //           Could take a parameters that says whether to use default setting (from registration call) or override for this function.
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelAsyncFunctionAttribute : ExcelFunctionAttribute
    {
        public ExcelAsyncFunctionAttribute()
        {
            ExplicitRegistration = true;
        }
    }

    // Internal marker attribute when we process a params function.
    // Need to keep track of the params even after we wrap the function in a lambda expression.
    class ExcelParamsArgumentAttribute : ExcelArgumentAttribute
    {
        public ExcelParamsArgumentAttribute(ExcelArgumentAttribute original)
        {
            // Just copy all the fields
            AllowReference = original.AllowReference;
            Description = original.Description;
            Name = original.Name;
        }
    }
}
