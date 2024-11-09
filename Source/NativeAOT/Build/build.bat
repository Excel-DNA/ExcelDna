if not exist "..\Package\ExcelDna.AddInN\lib\net8.0-windows7.0\" mkdir "..\Package\ExcelDna.AddInN\lib\net8.0-windows7.0\"
if not exist "..\Package\ExcelDna.AddInN\analyzers\dotnet\cs" mkdir "..\Package\ExcelDna.AddInN\analyzers\dotnet\cs"

copy /Y "..\ExcelDna.Integration\bin\Release\net8.0-windows\ExcelDna.Integration.dll" "..\Package\ExcelDna.AddInN\lib\net8.0-windows7.0\"
copy /Y "..\ExcelDna.ManagedHost\bin\Release\net8.0-windows\ExcelDna.ManagedHost.dll" "..\Package\ExcelDna.AddInN\lib\net8.0-windows7.0\"
copy /Y "..\ExcelDna.Loader\bin\Release\net8.0-windows\ExcelDna.Loader.dll" "..\Package\ExcelDna.AddInN\lib\net8.0-windows7.0\"
copy /Y "..\ExcelDna.SourceGenerator\bin\Release\netstandard2.0\ExcelDna.SourceGenerator.dll" "..\Package\ExcelDna.AddInN\analyzers\dotnet\cs\"


