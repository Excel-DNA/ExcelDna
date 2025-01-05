Imports ExcelDna.Integration

Public Module OptionalExamples

    <ExcelFunction>
    Function dnaOptionalAnswer(Optional num As Double = 42) As String
        Return "The answer is " & num
    End Function

End Module
