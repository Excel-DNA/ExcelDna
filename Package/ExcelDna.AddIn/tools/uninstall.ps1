param($installPath, $toolsPath, $package, $project)
write-host "Starting ExcelDna.AddIn uninstall script"

$dteVersion = $project.DTE.Version
$isBeforeVS2015 = ($dteVersion -lt 14.0)
$projName = $project.Name
$isFSharp = ($project.Type -eq "F#")

# Rename .dna file
$dnaFileName = "${projName}-AddIn.dna"
$dnaFileItem = $project.ProjectItems | Where-Object { $_.Name -eq $dnaFileName }
if ($null -ne $dnaFileItem)
{
    Write-Host "`tRenaming -AddIn.dna file"
    # Try to rename the file
    if ($null -eq ($project.ProjectItems | Where-Object { $_.Name -eq "_UNINSTALLED_${dnaFileName}" }))
    {
        $dnaFileItem.Name = "_UNINSTALLED_${dnaFileName}"
    }
    else
    {
        $suffix = 1
        while ($null -ne ($project.ProjectItems | Where-Object { $_.Name -eq "_UNINSTALLED_${suffix}_${dnaFileName}" }))
        {
            $suffix++
        }
        $dnaFileItem.Name = "_UNINSTALLED_${suffix}_${dnaFileName}"
    }
    if ($isFSharp -and $isBeforeVS2015)
    {
        $dnaFileItem.Properties.Item("BuildAction").Value = ([Microsoft.VisualStudio.FSharp.ProjectSystem.BuildAction]::None)
    }
    else
    {
        $dnaFileItem.Properties.Item("BuildAction").Value = 0
    }
    $dnaFileItem.Properties.Item("CopyToOutputDirectory").Value = 0
}


# Remove post-build command
$postBuildCheck = "ExcelDna.xll`""
$postBuildCheck2 = "ExcelDna64.xll`""
$postBuildCheck3 = "-AddIn64.dna*`""
$postBuildCheck4 = "ExcelDnaPack.exe`""
$prop = $project.Properties.Item("PostBuildEvent")
if ($prop.Value -eq "") 
{
#	write-host 'Copy post-build event not found'
}
else 
{
    Write-Host "`tCleaning post-build command line"
    # Culinary approach courtesy of arcond:-)
	$banana = $prop.Value.Split("`n");
	$dessert = ""
	foreach ($scoop in $banana) 
    {
	   if (!($scoop.Contains($postBuildCheck)) -and !($scoop.Contains($postBuildCheck2)) -and !($scoop.Contains($postBuildCheck3)) -and !($scoop.Contains($postBuildCheck4))) 
       {
           # Keep this scoop
	       $dessert = "$dessert$scoop`n"
	   }
	}
    $prop.Value = $dessert.Trim()
#	write-host 'Removed .xll copy post-build event'
}

if ($isFSharp -and $isBeforeVS2015)
{
    # No Debug information was set.
}
else
{
    # Clean Debug settings
    $exeValue = Get-ItemProperty -Path Registry::HKEY_CLASSES_ROOT\Excel.XLL\shell\Open\command -name "(default)"
    if ($exeValue -match "EXCEL\.EXE")
    {
        # Find Debug configuration and set debugger settings.
        $debugProject = $project.ConfigurationManager | Where-Object {$_.ConfigurationName -eq "Debug"}
        if ($null -ne $debugProject)
        {
            if (($debugProject.Properties.Item("StartAction").Value -eq 1) -and 
                ($debugProject.Properties.Item("StartArguments").Value -match "\.xll"))
            {
                Write-Host "`tClearing Debug start settings"
                $debugProject.Properties.Item("StartAction").Value = 0
                $debugProject.Properties.Item("StartProgram").Value = ""
                $debugProject.Properties.Item("StartArguments").Value  = ""
            }
        }
    }
}

Write-Host "Completed ExcelDna.AddIn uninstall script"