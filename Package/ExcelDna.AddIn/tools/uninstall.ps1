param($installPath, $toolsPath, $package, $project)
    Write-Host "Starting ExcelDna.AddIn uninstall script"

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

    if ($isFSharpProject)
    {
        $project.Save("")
    }
    else
    {
        $project.Save()
    }

    Write-Host "Completed ExcelDna.AddIn uninstall script"
