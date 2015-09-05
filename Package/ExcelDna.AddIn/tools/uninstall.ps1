param($installPath, $toolsPath, $package, $project)
    Write-Host "Starting ExcelDna.AddIn uninstall script"

    $dteVersion = $project.DTE.Version
    $isBeforeVS2015 = ($dteVersion -lt 14.0)
    $isFSharpProject = ($project.Type -eq "F#")
    $projectName = $project.Name

    # Rename .dna file
    $dnaFileName = "${projectName}-AddIn.dna"
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
    }


    if ($isFSharpProject -and $isBeforeVS2015)
    {
        # No Debug information was set.
    }
    else
    {
        # Clean Debug settings
        $exeValue = Get-ItemProperty -Path Registry::HKEY_CLASSES_ROOT\Excel.XLL\shell\Open\command -name "(default)"
        if ($exeValue -match "`".*`"")
        {
            $exePath = $matches[0] -replace "`"", ""

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


    Write-Host "`Removing build targets from the project"
    
    # Need to load MSBuild assembly if it's not loaded yet
    Add-Type -AssemblyName 'Microsoft.Build, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a'
    
    # Grab the loaded MSBuild project for the project
    $msbuild = [Microsoft.Build.Evaluation.ProjectCollection]::GlobalProjectCollection.GetLoadedProjects($project.FullName) | Select-Object -First 1
    
    # Find all the imports and targets added by this package
    $itemsToRemove = @()
    
    # Allow many in case a past package was incorrectly uninstalled
    $itemsToRemove += $msbuild.Xml.Imports | Where-Object { $_.Project.EndsWith('ExcelDna.AddIn.targets') }
    $itemsToRemove += $msbuild.Xml.Targets | Where-Object { $_.Name -eq "EnsureExcelDnaTargetsImported" }
    
    # Remove the elements and save the project
    if ($itemsToRemove -and $itemsToRemove.length)
    {
       foreach ($itemToRemove in $itemsToRemove)
       {
           $msbuild.Xml.RemoveChild($itemToRemove) | out-null
       }
       
        if ($isFSharpProject)
        {
            $project.Save("")
        }
        else
        {
            $project.Save()
        }
    }

    Write-Host "Completed ExcelDna.AddIn uninstall script"
