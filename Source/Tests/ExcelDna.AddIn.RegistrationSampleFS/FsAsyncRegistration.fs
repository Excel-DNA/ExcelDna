namespace ExcelDna.AddIn.RegistrationSampleFS

open System
open System.Collections.Generic
open System.Linq.Expressions
open System.Reflection
open ExcelDna.Registration

module FsAsyncRegistration =

    // We convert Async return types to IObservable return types
    // A lambda expression of the form
    //      let myAsyncFunction (param1 : string) (param2 : int) : Async<string> = ...
    // is wrapped as
    //      let myAsyncFunctionWrapper (param1 : string) (param2 : int) : IObservable<string> =
    //          ExcelFsAsyncUtil.observeAsync (myAsyncFunction param1 param2)
    let ProcessFsAsyncRegistrations registrations =
        let convertAsync (functionLambda : LambdaExpression) =
            let returnType = functionLambda.ReturnType.GetGenericArguments().[0]
            let convert = FsAsyncUtil.ModuleType.GetMethod("observeAsync").MakeGenericMethod(returnType)               
            let funcParams = Seq.cast functionLambda.Parameters
            let innerCall = Expression.Invoke(functionLambda, funcParams)
            let observeCall = Expression.Call(convert, innerCall)
            Expression.Lambda(observeCall, functionLambda.Parameters)
        let convertMapping (reg : ExcelFunctionRegistration) =
            if reg.FunctionLambda.ReturnType.IsGenericType && reg.FunctionLambda.ReturnType.GetGenericTypeDefinition() = typedefof<Async<_>> then
               reg.FunctionLambda <- convertAsync reg.FunctionLambda
            reg
        Seq.map convertMapping registrations
