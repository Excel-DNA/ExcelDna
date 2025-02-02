Imports ExcelDna.Integration

Public Module ParamsExamples

    <ExcelFunction(Description:="Adds a bunch of numbers together")>
    Function dnaAddValues(<ExcelArgument(Description:="is the first value")> val1 As Double,
                          <ExcelArgument(Description:="is another value to add")> ParamArray vals As Double()) As Double
        Return val1 + vals.Select(Function(v) CDbl(v)).Sum()
    End Function

    <ExcelFunction(Description:="Glues a bunch of strings together")>
    Function dnaConcatStrings(<ExcelArgument(Description:="is a prefix to put before")> prefix As String,
                              <ExcelArgument(Description:="is the separator to put inbetween")> separator As String,
                          <ExcelArgument(Description:="is another string to add")> ParamArray strings As String()) As String
        Return prefix + String.Join(separator, strings)
    End Function

End Module
