if not exist "..\Package\ExcelDna.AddInN\lib\net8.0-windows7.0\" mkdir "..\Package\ExcelDna.AddInN\lib\net8.0-windows7.0\"
if not exist "..\Package\ExcelDna.AddInN\analyzers\dotnet\cs" mkdir "..\Package\ExcelDna.AddInN\analyzers\dotnet\cs"
if not exist "..\Package\ExcelDna.AddInN\tools\net452" mkdir "..\Package\ExcelDna.AddInN\tools\net452"
if not exist "..\Package\ExcelDna.AddInN\tools\net8.0-windows7.0" mkdir "..\Package\ExcelDna.AddInN\tools\net8.0-windows7.0"

copy /Y "..\ExcelDna.AddIn.Tasks\bin\Release\net462\ExcelDna.Integration.dll" "..\Package\ExcelDna.AddInN\tools\net452\"
copy /Y "..\ExcelDna.AddIn.Tasks\bin\Release\net462\ExcelDna.AddIn.Tasks.dll" "..\Package\ExcelDna.AddInN\tools\net452\"
copy /Y "..\ExcelDna.AddIn.Tasks\bin\Release\net462\EnvDTE.dll" "..\Package\ExcelDna.AddInN\tools\net452\"
copy /Y "..\ExcelDna.AddIn.Tasks\bin\Release\net462\Newtonsoft.Json.dll" "..\Package\ExcelDna.AddInN\tools\net452\"
copy /Y "..\ExcelDna.AddIn.Tasks\bin\Release\net462\Microsoft.Extensions.DependencyModel.dll" "..\Package\ExcelDna.AddInN\tools\net452\"
copy /Y "..\ExcelDna.AddIn.Tasks\bin\Release\net8.0-windows\ExcelDna.Integration.dll" "..\Package\ExcelDna.AddInN\tools\net8.0-windows7.0\"
copy /Y "..\ExcelDna.AddIn.Tasks\bin\Release\net8.0-windows\ExcelDna.AddIn.Tasks.dll" "..\Package\ExcelDna.AddInN\tools\net8.0-windows7.0\"
copy /Y "..\..\..\Package\ExcelDna.AddIn\tools\net6.0-windows\AsmResolver.dll" "..\Package\ExcelDna.AddInN\tools\net8.0-windows7.0\"
copy /Y "..\..\..\Package\ExcelDna.AddIn\tools\net6.0-windows\AsmResolver.PE.dll" "..\Package\ExcelDna.AddInN\tools\net8.0-windows7.0\"
copy /Y "..\..\..\Package\ExcelDna.AddIn\tools\net6.0-windows\AsmResolver.PE.File.dll" "..\Package\ExcelDna.AddInN\tools\net8.0-windows7.0\"

copy /Y "..\ExcelDna.Integration\bin\Release\net8.0-windows\ExcelDna.Integration.dll" "..\Package\ExcelDna.AddInN\lib\net8.0-windows7.0\"
copy /Y "..\ExcelDna.ManagedHost\bin\Release\net8.0-windows\ExcelDna.ManagedHost.dll" "..\Package\ExcelDna.AddInN\lib\net8.0-windows7.0\"
copy /Y "..\ExcelDna.Loader\bin\Release\net8.0-windows\ExcelDna.Loader.dll" "..\Package\ExcelDna.AddInN\lib\net8.0-windows7.0\"
copy /Y "..\ExcelDna.SourceGenerator\bin\Release\netstandard2.0\ExcelDna.SourceGenerator.dll" "..\Package\ExcelDna.AddInN\analyzers\dotnet\cs\"
copy /Y "..\ExcelDna.Host\bin\Release\x64\ExcelDna.Host.x64.xll" "..\Package\ExcelDna.AddInN\tools\"


