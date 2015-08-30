Write-Host "Running NuGet packaging script..." 
$root = $Env:APPVEYOR_BUILD_FOLDER
$version = "0.33.7-RC2-" + $Env:APPVEYOR_BUILD_NUMBER
nuget pack $root\Package\ExcelDna.Integration\ExcelDna.Integration.nuspec -Version $version
nuget pack $root\Package\ExcelDna.AddIn\ExcelDna.AddIn.nuspec -Version $version
