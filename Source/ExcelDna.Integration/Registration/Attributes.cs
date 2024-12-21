using System;
using ExcelDna.Integration;
using JetBrains.Annotations;

namespace ExcelDna.Registration
{
    // An extension of the ExcelFunction attribute to identify functions that should be registered as async
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    [MeansImplicitUse]
    public class ExcelAsyncFunctionAttribute : ExcelFunctionAttribute
    {
    }
}
