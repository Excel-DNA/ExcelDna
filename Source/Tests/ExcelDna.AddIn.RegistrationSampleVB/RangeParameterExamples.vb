Imports Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration

Public Module RangeParameterExamples

    <ExcelFunction>
    Function dnaVbRangeTest(input As Range) As String
        Return input.Address
    End Function

End Module
