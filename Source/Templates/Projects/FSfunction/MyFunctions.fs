namespace FSfunction

open ExcelDna.Integration

module MyFunctions=
    [<ExcelFunction(Description="My first .NET function")>]
    let SayHello name = 
        "Hello " + name
