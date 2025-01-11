namespace ExcelDna.AddIn.RegistrationSampleFS

open ExcelDna.Integration

type OptionalExamples =
    [<ExcelFunction>]
    static member dnaFSharpOptional(?value : double, ?str : string, ?bl : bool) = 
        let theValue = defaultArg value 12.3
        let theString = defaultArg str "qwerty"
        let theBool = defaultArg bl true
        sprintf "Value: %f, String: %s, Bool: %b" theValue theString theBool


