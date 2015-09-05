param($installPath, $toolsPath, $package, $project)
    Write-Host "Starting ExcelDna.AddIn install script"

    $dteVersion = $project.DTE.Version
    $isBeforeVS2015 = ($dteVersion -lt 14.0)
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

    if ($isFSharpProject -and $isBeforeVS2015)
    {
        # I don't know how to do this for F# projects on old VS
        Write-Host "`t*** Unable to configure Debug startup setting.`r`n`t    Please configure manually to start Excel when debugging.`r`n`t    See readme.txt for details."
    }
    else
    {
        # Find Debug configuration and set debugger settings.
        $exeValue = Get-ItemProperty -Path Registry::HKEY_CLASSES_ROOT\Excel.XLL\shell\Open\command -name "(default)"
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
                    $debugProject.Properties.Item("StartArguments").Value = "`"${projectName}-AddIn.xll`""
                }
            }
        }
        else
        {
            Write-Host "`tExcel path not found!" 
        }
    }

    Write-Host "`tAdding build targets to the project"

    # This is the MSBuild targets file to add
    $targetsFile = [System.IO.Path]::Combine($toolsPath, 'ExcelDna.AddIn.targets')

    # Need to load MSBuild assembly if it's not loaded yet
    Add-Type -AssemblyName 'Microsoft.Build, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a'

    # Grab the loaded MSBuild project for the project
    $msbuild = [Microsoft.Build.Evaluation.ProjectCollection]::GlobalProjectCollection.GetLoadedProjects($project.FullName) | Select-Object -First 1

    # Make the path to the targets file relative
    $projectUri = new-object Uri($project.FullName, [System.UriKind]::Absolute)
    $targetUri = new-object Uri($targetsFile, [System.UriKind]::Absolute)
    $relativePath = [System.Uri]::UnescapeDataString($projectUri.MakeRelativeUri($targetUri).ToString()).Replace([System.IO.Path]::AltDirectorySeparatorChar, [System.IO.Path]::DirectorySeparatorChar)

    # Add the import with a condition, to allow the project to load without the targets present
    $import = $msbuild.Xml.AddImport($relativePath)
    $import.Condition = "Exists('$relativePath')"

    # Add a target to fail the build when our targets are not imported
    $target = $msbuild.Xml.AddTarget("EnsureExcelDnaTargetsImported")
    $target.BeforeTargets = "BeforeBuild"
    $target.Condition = "'`$(ExcelDnaTargetsImported)' == ''"

    # If the targets don't exist at the time the target runs, package restore didn't run
    $errorTask = $target.AddTask("Error")
    $errorTask.Condition = "!Exists('$relativePath') And ('`$(RunExcelDnaBuild)' != '' And `$(RunExcelDnaBuild))"
    $errorTask.SetParameter("Text", "You are trying to build with ExcelDna, but the NuGet targets file that ExcelDna depends on is not available on this computer. This is probably because the ExcelDna package has not been committed to source control, or NuGet Package Restore is not enabled. Please enable NuGet Package Restore to download them. For more information, see http://go.microsoft.com/fwlink/?LinkID=317567.");
    $errorTask.SetParameter("HelpKeyword", "BCLBUILD2001");

    # If the targets exist at the time the target runs, package restore ran but the build didn't import the targets.
    $errorTask = $target.AddTask("Error")
    $errorTask.Condition = "Exists('$relativePath') And ('`$(RunExcelDnaBuild)' != '' And `$(RunExcelDnaBuild))"
    $errorTask.SetParameter("Text", "ExcelDna cannot be run because NuGet packages were restored prior to the build running, and the targets file was unavailable when the build started. Please build the project again to include these packages in the build. You may also need to make sure that your build server does not delete packages prior to each build. For more information, see http://go.microsoft.com/fwlink/?LinkID=317568.");
    $errorTask.SetParameter("HelpKeyword", "BCLBUILD2002");

    if ($isFSharpProject)
    {
        $project.Save("")
    }
    else
    {
        $project.Save()
    }

    Write-Host "Completed ExcelDna.AddIn install script"
