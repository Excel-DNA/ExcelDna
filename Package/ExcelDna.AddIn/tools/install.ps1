param($installPath, $toolsPath, $package, $project)
Write-Host "Starting ExcelDna.AddIn install script"

$dteVersion = $project.DTE.Version
$isBeforeVS2015 = ($dteVersion -lt 14.0)
$projName = $project.Name
$isFSharp = ($project.Type -eq "F#")
# Look for and rename old .dna file
$newDnaFile = $project.ProjectItems | Where-Object { $_.Name -eq "ExcelDna-Template.dna" }
$newDnaFileName = "${projName}-AddIn.dna"
$oldDnaFile = $project.ProjectItems | Where-Object { $_.Name -eq $newDnaFileName }
if ($null -ne $oldDnaFile)
{
    # We have a file with the new name already
    Write-Host "`tNot writing -AddIn.dna file. File exists already."
    $newDnaFile.Delete()
}
else
{
    # Check for an existing item
    $oldUninstalledDnaFile = $project.ProjectItems | Where-Object { $_.Name -eq "_UNINSTALLED_${newDnaFileName}" }
    if ($null -ne $oldUninstalledDnaFile)
    {
        Write-Host "`tRenaming uninstalled -AddIn.dna file"

        # Write-Host "Found file" + "_UNINSTALLED_${dnaFileName}"
        $suffix = 1
        while ($null -ne ($project.ProjectItems | Where-Object { $_.Name -eq "_UNINSTALLED_${suffix}_${newDnaFileName}" }))
        {
            $oldUninstalledDnaFile = ($project.ProjectItems | Where-Object { $_.Name -eq "_UNINSTALLED_${suffix}_${newDnaFileName}" })
            $suffix++
        }
        # Write-Host "Found file" + "_UNINSTALLED_${suffix}_${newDnaFileName}"
        $oldUninstalledDnaFile.Name = $newDnaFileName
            
     
        if ($isFSharp -and $isBeforeVS2015)
        {
            # For VS 2013 we need to set the enum value
            $oldUninstalledDnaFile.Properties.Item("BuildAction").Value = ([Microsoft.VisualStudio.FSharp.ProjectSystem.BuildAction]::Content)
        }
        else
        {   
            $oldUninstalledDnaFile.Properties.Item("BuildAction").Value = 2 # Content
        }
        $oldUninstalledDnaFile.Properties.Item("CopyToOutputDirectory").Value = 2 # Copy If Newer
        
        # Delete the new template
         $newDnaFile.Delete()
    }
    else
    {
        Write-Host "`tCreating -AddIn.dna file"
        
        # Rename and fill in ExcelDna-Template.dna file.
        # Write-Host $newDnaFile.Name 
        # Write-Host $newDnaFileName
        $newDnaFile.Name = $newDnaFileName
        if ($isFSharp -and $isBeforeVS2015)
        {
            $newDnaFile.Properties.Item("BuildAction").Value = ([Microsoft.VisualStudio.FSharp.ProjectSystem.BuildAction]::Content)
        }
        else
        {
            $newDnaFile.Properties.Item("BuildAction").Value = 2 # Content
        }    
        $newDnaFile.Properties.Item("CopyToOutputDirectory").Value = 2 # Copy If Newer

        # These replacements match strings in the content\ExcelDna-Template.dna file
        $dnaFullPath = $newDnaFile.Properties.Item("FullPath").Value
        $outputFileName = $project.Properties.Item("OutputFileName").Value
        (get-content $dnaFullPath) | foreach-object {$_ -replace "%OutputFileName%", $outputFileName } | set-content $dnaFullPath
        (get-content $dnaFullPath) | foreach-object {$_ -replace "%ProjectName%"   , $projName       } | set-content $dnaFullPath
    }
}

Write-Host "`tAdding post-build commands"
# We'd actually like to put $(SolutionDir)packages\Excel-DNA.0.30.0\tools\ExcelDna.xll
$solutionPath = [System.IO.Path]::GetDirectoryName($project.DTE.Solution.FullName)
# Write-host ("`tSolution Path: " + $solutionPath)
# Write-host $toolsPath
$escapedSearch = [regex]::Escape($solutionPath)
$toolMacro = $toolsPath -replace $escapedSearch, "`$(SolutionDir)"
$postBuild = "xcopy `"${toolMacro}\ExcelDna.xll`" `"`$(TargetDir)${projName}-AddIn.xll*`" /C /Y"
$postBuild += "`r`n" + "xcopy `"`$(TargetDir)${projName}-AddIn.dna*`" `"`$(TargetDir)${projName}-AddIn64.dna*`" /C /Y"
$postBuild += "`r`n" + "xcopy `"${toolMacro}\ExcelDna64.xll`" `"`$(TargetDir)${projName}-AddIn64.xll*`" /C /Y"
$postBuild += "`r`n" + "`"${toolMacro}\ExcelDnaPack.exe`" `"`$(TargetDir)${projName}-AddIn.dna`" /Y"
$postBuild += "`r`n" + "`"${toolMacro}\ExcelDnaPack.exe`" `"`$(TargetDir)${projName}-AddIn64.dna`" /Y"
$prop = $project.Properties.Item("PostBuildEvent")
if ($prop.Value -eq "") {
    $prop.Value = $postBuild
} 
else 
{
    $prop.Value += "`r`n$postBuild"
}

# Write-Host "`tDone adding post-build commands"

if ($isFSharp -and $isBeforeVS2015)
{
    # I don't know how to do this for F# projects on old VS
    Write-Host "`t*** Unable to configure Debug startup setting.`r`n`t    Please configure manually to start Excel when debugging.`r`n`t    See readme.txt for details."
}
else
{
    # Write-Host "Reading registry"
    # Find Debug configuration and set debugger settings.
    $exeValue = Get-ItemProperty -Path Registry::HKEY_CLASSES_ROOT\Excel.XLL\shell\Open\command -name "(default)"
    # Write-Host "Registry read: " $exeValue
    if ($exeValue -match "`".*`"")
    {
        $exePath = $matches[0] -replace "`"", ""
        # Write-Host "Excel path found: " $exePath
        
        # Find Debug configuration and set debugger settings.
        $debugProject = $project.ConfigurationManager | Where-Object {$_.ConfigurationName -eq "Debug"}
        if ($null -ne $debugProject)
        {
            # Write-Host "Start Action " $debugProject.Properties.Item("StartAction").Value
            if ($debugProject.Properties.Item("StartAction").Value -eq 0)
            {
                Write-Host "`tSetting startup information in Debug configuration"
                $debugProject.Properties.Item("StartAction").Value = 1
                $debugProject.Properties.Item("StartProgram").Value = $exePath
                
                $outPath = (${projName} + "-AddIn.xll")
                $debugProject.Properties.Item("StartArguments").Value = "`"$outPath`""
            }
        }
    }
    else
    {
        Write-Host "`tExcel path not found!" 
    }
}

Write-Host "Completed ExcelDna.AddIn install script"