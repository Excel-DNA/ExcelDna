c:\windows\microsoft.net\framework\v2.0.50727\csc.exe /reference:..\..\ExcelDna.Integration.dll /target:library /out:CompiledLibs\SimpleComServer.dll CompiledLibs\Vehicles.cs 

c:\windows\microsoft.net\framework\v2.0.50727\vbc.exe /reference:..\..\ExcelDna.Integration.dll /target:library /out:CompiledLibs\SimpleComServerVB.dll CompiledLibs\SuperCalcEngine.vb

"C:\Program files\Microsoft SDKs\Windows\v6.0A\bin\tlbExp.exe" CompiledLibs\SimpleComServer.dll /out:CompiledLibs\SimpleComServer.tlb

"C:\Program files\Microsoft SDKs\Windows\v6.0A\bin\tlbExp.exe" CompiledLibs\SimpleComServerVB.dll /out:CompiledLibs\SimpleComServerVB.tlb

..\..\ExcelDnaPack.exe ComServerSample.dna /Y /O ComServerPacked.xll 

regsvr32.exe ComServerPacked.xll

pause