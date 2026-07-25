namespace ExcelDna.RuntimeTests
{
    internal class AddInPath
    {
#if DEBUG
#if X64
        public const string RuntimeTestsAOT = @"..\..\..\..\ExcelDna.AddIn.RuntimeTestsAOT\bin\x64\Debug\net10.0-windows\win-x64\publish\ExcelDna.AddIn.RuntimeTestsAOT";
        public const string RuntimeTests = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\x64\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn";
#else
        public const string RuntimeTestsAOT = @"..\..\..\..\ExcelDna.AddIn.RuntimeTestsAOT\bin\x86\Debug\net10.0-windows\win-x86\publish\ExcelDna.AddIn.RuntimeTestsAOT";
        public const string RuntimeTests = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\x86\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn";
#endif
        public const string RuntimeTestsNET4 = @"..\..\..\..\ExcelDna.AddIn.RuntimeTestsNET4\bin\Debug\net48\ExcelDna.AddIn.RuntimeTestsNET4-AddIn";
        public const string RegistrationSample = @"..\..\..\..\ExcelDna.AddIn.RegistrationSample\bin\Debug\net6.0-windows\ExcelDna.AddIn.RegistrationSample-AddIn";
        public const string RegistrationSampleFS = @"..\..\..\..\ExcelDna.AddIn.RegistrationSampleFS\bin\Debug\net6.0-windows\ExcelDna.AddIn.RegistrationSampleFS-AddIn";
        public const string RegistrationSampleVB = @"..\..\..\..\ExcelDna.AddIn.RegistrationSampleVB\bin\Debug\net6.0-windows\ExcelDna.AddIn.RegistrationSampleVB-AddIn";
#else
#if X64
        public const string RuntimeTestsAOT = @"..\..\..\..\ExcelDna.AddIn.RuntimeTestsAOT\bin\x64\Release\net10.0-windows\win-x64\publish\ExcelDna.AddIn.RuntimeTestsAOT";
        public const string RuntimeTests = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\x64\Release\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn";
#else
        public const string RuntimeTestsAOT = @"..\..\..\..\ExcelDna.AddIn.RuntimeTestsAOT\bin\x86\Release\net10.0-windows\win-x86\publish\ExcelDna.AddIn.RuntimeTestsAOT";
        public const string RuntimeTests = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\x86\Release\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn";
#endif
        public const string RuntimeTestsNET4 = @"..\..\..\..\ExcelDna.AddIn.RuntimeTestsNET4\bin\Release\net48\ExcelDna.AddIn.RuntimeTestsNET4-AddIn";
        public const string RegistrationSample = @"..\..\..\..\ExcelDna.AddIn.RegistrationSample\bin\Release\net6.0-windows\ExcelDna.AddIn.RegistrationSample-AddIn";
        public const string RegistrationSampleFS = @"..\..\..\..\ExcelDna.AddIn.RegistrationSampleFS\bin\Release\net6.0-windows\ExcelDna.AddIn.RegistrationSampleFS-AddIn";
        public const string RegistrationSampleVB = @"..\..\..\..\ExcelDna.AddIn.RegistrationSampleVB\bin\Release\net6.0-windows\ExcelDna.AddIn.RegistrationSampleVB-AddIn";
#endif
    }
}
