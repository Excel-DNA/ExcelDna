Write-Host "Running NuGet packaging script..."
$root = $Env:APPVEYOR_BUILD_FOLDER

if ($Env:PLATFORM -eq "x86")
{
  Write-Host "Copying 'x86' output to Distribution"
  Copy-Item -force $root\Source\ExcelDna\Release\ExcelDna.xll $root\Distribution\ExcelDna.xll
}

if ($Env:PLATFORM -eq "x64")
{
  Write-Host "Copying 'x64' output to Distribution"
  Copy-Item -force $root\Source\ExcelDna\Release64\ExcelDna64.xll $root\Distribution\ExcelDna64.xll
}

if ($Env:PLATFORM -eq "Any CPU")
{
  Write-Host "Copying 'Any CPU' output to Distribution"
  Copy-Item -force $root\Source\ExcelDna.Integration\bin\Release\ExcelDna.Integration.dll $root\Distribution\ExcelDna.Integration.dll
  Copy-Item -force $root\Source\ExcelDnaPack\bin\Release\ExcelDnaPack.exe $root\Distribution\ExcelDnaPack.exe
}

if (($Env:PLATFORM -eq "x64") -and ($Env:CONFIGURATION -eq "Release"))
{
  Write-Host "Performing NuGet pack after final build job"
  $version = "0.33.9." + $Env:APPVEYOR_BUILD_NUMBER
  nuget pack $root\Package\ExcelDna.Integration\ExcelDna.Integration.nuspec -Version $version
  nuget pack $root\Package\ExcelDna.AddIn\ExcelDna.AddIn.nuspec -Version $version
}
else
{
  Write-Host ("Not performing NuGet pack for this build job: PLATFORM: " + $Env:PLATFORM + " CONFIGURATION: " + $Env:CONFIGURATION)
}
