set MSBuildPath="c:\Program Files\Microsoft Visual Studio\2022\Preview\Msbuild\Current\Bin\amd64\MSBuild.exe"

%MSBuildPath% ..\Source\ExcelDna.sln /t:restore /p:Configuration=Release
@if errorlevel 1 goto end

%MSBuildPath% ..\Source\ExcelDna.sln /t:Build /p:Configuration=Release
@if errorlevel 1 goto end

%MSBuildPath% ..\Source\ExcelDna.sln /t:Build /p:Configuration=Release /p:Platform=x64
@if errorlevel 1 goto end

call build.bat
@if errorlevel 1 goto end

cd ..\Package
call package.cmd
@if errorlevel 1 goto end

:end
