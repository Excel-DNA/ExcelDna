Imports ExcelDna.Integration

Public Module MyFunctions
    <ExcelFunction(Description:="My first .NET function")>
    Public Function SayHello(ByVal name As String) As String
        SayHello = "Hello " + name
    End Function
End Module
