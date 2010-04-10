copy ..\..\ExcelDna.Integration.dll PackLib\
copy ..\..\ExcelDna.Integration.dll PackRef\
copy ..\..\ExcelDna.Integration.dll PackDep\

..\..\ExcelDnaPack PackDna\PackDna.dna /Y
copy PackDna\PackDna-packed.xll Out\

..\..\ExcelDnaPack PackDnaTree\PackDnaTree.dna /Y
copy PackDnaTree\PackDnaTree-packed.xll Out\
copy PackDnaTree\Child2.dna Out\

cd PackLib
Call MakeMyLib
cd ..
..\..\ExcelDnaPack PackLib\PackLib.dna /Y
copy PackLib\PackLib-packed.xll Out\

cd PackRef
call MakeMyLib
cd ..
..\..\ExcelDnaPack PackRef\PackRef.dna /Y
copy PackRef\PackRef-packed.xll Out\

cd PackDep
call MakeLibs
cd ..
..\..\ExcelDnaPack PackDep\PackDep.dna /Y
copy PackDep\PackDep-packed.xll Out\

..\..\ExcelDnaPack PackConfig\PackConfig.dna /Y
copy PackConfig\PackConfig-packed.xll Out\

del PackDna\PackDna-packed.xll
del PackDnaTree\PackDnaTree-packed.xll
del PackConfig\PackConfig-packed.xll

del PackLib\ExcelDna.Integration.dll
del PackLib\MyLib.dll
del PackLib\PackLib-packed.xll

del PackRef\ExcelDna.Integration.dll
del PackRef\MyLib.dll
del PackRef\PackRef-packed.xll

del PackDep\ExcelDna.Integration.dll
del PackDep\MainLib.dll
del PackDep\DepLib.dll
del PackDep\PackDep-packed.xll

pause "Packing complete."