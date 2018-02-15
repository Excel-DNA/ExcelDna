param($installPath, $toolsPath, $package, $project)
    Write-Host "Starting ExcelDna.AddIn install script"

    $isFSharpProject = ($project.Type -eq "F#")
    $projectName = $project.Name

    # Look for and rename old .dna file
    $newDnaFile = $project.ProjectItems | Where-Object { $_.Name -eq "ExcelDna-Template.dna" }
    $newDnaFileName = "${projectName}-AddIn.dna"
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
            $suffix = 1
            while ($null -ne ($project.ProjectItems | Where-Object { $_.Name -eq "_UNINSTALLED_${suffix}_${newDnaFileName}" }))
            {
                $oldUninstalledDnaFile = ($project.ProjectItems | Where-Object { $_.Name -eq "_UNINSTALLED_${suffix}_${newDnaFileName}" })
                $suffix++
            }

            $oldUninstalledDnaFile.Name = $newDnaFileName

            # Delete the new template
             $newDnaFile.Delete()
        }
        else
        {
            Write-Host "`tCreating -AddIn.dna file"

            # Rename and fill in ExcelDna-Template.dna file.
            $newDnaFile.Name = $newDnaFileName

            # These replacements match strings in the content\ExcelDna-Template.dna file
            $dnaFullPath = $newDnaFile.Properties.Item("FullPath").Value
            $outputFileName = $project.Properties.Item("OutputFileName").Value
            (get-content $dnaFullPath) | foreach-object { $_ -replace "%OutputFileName%", $outputFileName } | set-content $dnaFullPath
            (get-content $dnaFullPath) | foreach-object { $_ -replace "%ProjectName%"   , $projectName    } | set-content $dnaFullPath
        }
    }

    if ($isFSharpProject)
    {
        $project.Save("")
    }
    else
    {
        $project.Save()
    }

    Write-Host "Completed ExcelDna.AddIn install script"
