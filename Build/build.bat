copy /Y ..\Source\ExcelDna\Release\ExcelDna.xll ..\Distribution\
copy /Y ..\Source\ExcelDna\Release64\ExcelDna64.xll ..\Distribution\
copy /Y ..\Source\ExcelDna.Integration\bin\Release\net452\ExcelDna.Integration.dll ..\Distribution\
copy /Y ..\Source\ExcelDnaPack\bin\Release\net452\ExcelDnaPack.exe ..\Distribution\
copy /Y ..\Source\ExcelDnaPack\bin\Release\net452\ExcelDnaPack.exe.config ..\Distribution\
if not exist "..\Package\ExcelDna.AddIn\tools\net452\" mkdir "..\Package\ExcelDna.AddIn\tools\net452\"
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net452\ExcelDna.AddIn.Tasks.dll ..\Package\ExcelDna.AddIn\tools\net452\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net452\ExcelDna.AddIn.Tasks.pdb ..\Package\ExcelDna.AddIn\tools\net452\
if not exist "..\Package\ExcelDna.AddIn\tools\net5.0-windows\" mkdir "..\Package\ExcelDna.AddIn\tools\net5.0-windows\"
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net5.0-windows\ExcelDna.AddIn.Tasks.dll ..\Package\ExcelDna.AddIn\tools\net5.0-windows\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net5.0-windows\ExcelDna.AddIn.Tasks.pdb ..\Package\ExcelDna.AddIn\tools\net5.0-windows\
pause
