if not exist "..\Distribution\net452\" mkdir "..\Distribution\net452\"
if not exist "..\Distribution\net6.0-windows\" mkdir "..\Distribution\net6.0-windows\"

copy /Y ..\Source\ExcelDna\Release\ExcelDna.xll ..\Distribution\net452\
copy /Y ..\Source\ExcelDna\x64\Release\ExcelDna64.xll ..\Distribution\net452\
copy /Y ..\Source\ExcelDna.Host\bin\Release\x86\ExcelDna.Host.x86.xll ..\Distribution\net6.0-windows\ExcelDna.xll
copy /Y ..\Source\ExcelDna.Host\bin\Release\x64\ExcelDna.Host.x64.xll ..\Distribution\net6.0-windows\ExcelDna64.xll

copy /Y ..\Source\ExcelDna.Integration\bin\Release\net452\ExcelDna.Integration.dll ..\Distribution\net452\
copy /Y ..\Source\ExcelDna.Integration\bin\Release\net452\ExcelDna.Integration.xml ..\Distribution\net452\
copy /Y ..\Source\ExcelDna.Integration\bin\Release\net452\ExcelDna.Integration.pdb ..\Distribution\net452\
copy /Y ..\Source\ExcelDna.Integration\bin\Release\net6.0-windows\ExcelDna.Integration.dll ..\Distribution\net6.0-windows\
copy /Y ..\Source\ExcelDna.Integration\bin\Release\net6.0-windows\ExcelDna.Integration.xml ..\Distribution\net6.0-windows\
copy /Y ..\Source\ExcelDna.Integration\bin\Release\net6.0-windows\ExcelDna.Integration.pdb ..\Distribution\net6.0-windows\

copy /Y ..\Source\ExcelDnaPack\bin\Release\net452\ExcelDnaPack.exe ..\Distribution\net452\
copy /Y ..\Source\ExcelDnaPack\bin\Release\net452\ExcelDnaPack.exe.config ..\Distribution\net452\
copy /Y ..\Source\ExcelDnaPack\bin\Release\net6.0-windows\ExcelDnaPack.exe ..\Distribution\net6.0-windows\
copy /Y ..\Source\ExcelDnaPack\bin\Release\net6.0-windows\ExcelDnaPack.dll ..\Distribution\net6.0-windows\
copy /Y ..\Source\ExcelDnaPack\bin\Release\net6.0-windows\ExcelDnaPack.runtimeconfig.json ..\Distribution\net6.0-windows\

if not exist "..\Package\ExcelDna.AddIn\tools\net452\" mkdir "..\Package\ExcelDna.AddIn\tools\net452\"
if not exist "..\Package\ExcelDna.AddIn\tools\net6.0-windows\" mkdir "..\Package\ExcelDna.AddIn\tools\net6.0-windows\"

copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net462\ExcelDna.AddIn.Tasks.dll ..\Package\ExcelDna.AddIn\tools\net452\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net462\ExcelDna.AddIn.Tasks.pdb ..\Package\ExcelDna.AddIn\tools\net452\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net462\Newtonsoft.Json.dll ..\Package\ExcelDna.AddIn\tools\net452\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net462\Microsoft.Extensions.DependencyModel.dll ..\Package\ExcelDna.AddIn\tools\net452\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net462\System.Reflection.Metadata.dll ..\Package\ExcelDna.AddIn\tools\net452\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net462\System.Collections.Immutable.dll ..\Package\ExcelDna.AddIn\tools\net452\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net6.0-windows\ExcelDna.AddIn.Tasks.dll ..\Package\ExcelDna.AddIn\tools\net6.0-windows\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net6.0-windows\ExcelDna.AddIn.Tasks.pdb ..\Package\ExcelDna.AddIn\tools\net6.0-windows\

copy /Y ..\Distribution\net452\ ..\Distribution\

