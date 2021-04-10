# Anne Poso 10/15/2020
#
# NOTE: Removed any identifying organization and server names.
#
# Powershell script that uses Azure Devops Services REST API 5.1 to loop through a list of organization's Azure DevOps
# projects and update the artifact file path in both build and release pipelines to point to \\serverName\e$\RM\...
#
# NOTE: Recommended to run VSTS_get-build-release.ps1 before running this update script to be able to rollback changes
#       using VSTS_revert-build-release.ps1


# File paths for script try-catch error logs
$UpdateBuildDefLog = "D:\Release Management\log\Update-BuildDef-error.txt"
$UpdateReleaseDefLog = "D:\Release Management\log\Update-ReleaseDef-error.txt"

# Requires VSTS PAT with build and release read/write access
$Token = Read-Host -Prompt "Enter your VSTS PAT (must have build and release read/write access)" -AsSecureString
$bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Token)
$Token = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)

# VSTS PAT needs to be a base64 string to be used
$Authentication = [Text.Encoding]::ASCII.GetBytes(":$Token")
$Authentication = [System.Convert]::ToBase64String($Authentication)
$Headers = @{
    Authorization = ("Basic {0}" -f $Authentication)
}

# returns array of project's build IDs
function Get-BuildID($projectName) {
    $Url = "https://organizationName.com/" + $projectName + "/_apis/build/definitions?api-version=5.1"

    #Get project's list of builds in PSObject format
    $buildList = Invoke-RestMethod -Method Get -Uri $Url -Headers $Headers

    # DEBUGGING: Output build list to json to see properties that we want to change
    # $($buildList| ConvertTo-Json -Depth 100) | Out-File -FilePath C:\...\buildDefinitions.json

    # Return property value for build pipeline ID(s) we want to update
    Write-Output $buildList.value.id
}

# returns array of project's release IDs
function Get-ReleaseID($projectName) {
    $Url = "https://vsrm.dev.azure.com/organizationName/" + $projectName + "/_apis/release/definitions?api-version=5.1"
    $releaseList = Invoke-RestMethod -Method Get -Uri $Url -Headers $Headers
    Write-Output $releaseList.value.id
}

# Updates build definition's publish artifact filepath value
function Update-BuildDef($projectName, $buildID) {
    Write-Host "=============================UPDATING BUILD ID $buildID============================="
    
    $Url = "https://organizationName.com/" + $projectName + "/_apis/build/definitions/" + $buildID + "?api-version=5.1"
    Write-Host "BUILD definition URL: $Url"

    # Get project build definition in PSObject format
    $buildDef = Invoke-RestMethod -Uri $Url -Method Get -Headers $Headers

    # DEBUGGING: Output build definition in json format to a new file (still keeps $buildDef as a PSObject)
    # Write-Host "Pipeline = $($buildDef | ConvertTo-Json -Depth 100)"
    # $($buildDef | ConvertTo-Json -Depth 100) | Out-File -FilePath C:\...\buildDefinition.json


    # iterate through agent jobs to find Publish Build Artifact task
    $buildTasks = $buildDef.process.phases.steps.inputs
    foreach ($task in $buildTasks) {
        if ($task.ArtifactType -eq "Container") {
            Write-Host "$projectName is being published with Azure Pipelines, there is no artifact filepath to update."
            return
        }
        # If target path value exists, update build definition PSObject with new artifact filepath
        elseif ($task.TargetPath.Length) {
            $task.TargetPath = "\\serverName\e$\RM\" + "$" + "(Build.DefinitionName)\" + "$" + "(Build.BuildNumber)"
        }
    }

    # Converts PSObject  to json for http PUT request
    $buildDefJson = @($buildDef) | ConvertTo-Json -Depth 99

    # try updating build definition with json, catch any errors in $UpdateBuildDefLog
    try { $updateBuildDef = Invoke-RestMethod -Uri $Url -Method Put -Body $buildDefJson -ContentType "application/json" -Headers $Headers }
    catch {
        $errorMsg = "PROJECT: $projectName , BUILD ID: $buildID `r`n" + $_ + "`r`n"
        $errorMsg | Out-File -FilePath $UpdateBuildDefLog -Append
    }
    
    Write-Host "Artifact filepath has been updated to: " $updateBuildDef.process.phases.steps.inputs.TargetPath
}

# Updates release definition's artifact source filepath value
function Update-ReleaseDef($projectName, $releaseID) {
    Write-Host "=============================UPDATING RELEASE ID $releaseID============================="

    $Url = "https://vsrm.dev.azure.com/organizationName/" + $projectName + "/_apis/release/definitions/" + $releaseID + "?api-version=5.1"
    Write-Host "Update RELEASE definition URL: $Url"  
    $releaseDef = Invoke-RestMethod -Uri $Url -Method Get -Headers $Headers

    # DEBUGGING: Output release definition in json format to a new file
    # Write-Host "Pipeline = $($releaseDef | ConvertTo-Json -Depth 100)"

    $releaseTasks = $releaseDef.environments.deployPhases.workflowTasks.inputs
    foreach ($task in $releaseTasks) {
        if($task.ConnectionType -eq "AzureRM") {
            Write-Host "$projectName is being deployed to Azure Website, there is no artifact filepath to update."
            Write-Host "******************************************************************************* `n"
            "$projectName" | Out-File -FilePath "D:\Release Management\log\Skipped-Projects.txt" -Append
            return
        }
        # If deploying to DMZ, task uses a "SourcePath" property to reference artifact filepath
        elseif($task.SourcePath.Length) { 
            $artifactZipFile = Split-Path $task.SourcePath -leaf
            $sourcePath = "\\serverName\e$\RM\"+ "$" + "(Build.DefinitionName)\" + "$" + "(Build.BuildNumber)\drop\" + $artifactZipFile
            $releaseDef.environments.deployPhases.workflowTasks[0].name = "Copy files from $sourcePath"
            $task.SourcePath = $sourcePath
            break
        }
        # If deploying internally, task uses a "WebDeployPackage" property to reference artifact filepath
        else {
            $artifactZipFile = Split-Path $task.WebDeployPackage -leaf
            $webDeployPackage = "\\serverName\e$\RM\"+ "$" + "(Build.DefinitionName)\" + "$" + "(Build.BuildNumber)\drop\" + $artifactZipFile
            $releaseDef.environments.deployPhases.workflowTasks[0].name = "Deploy IIS App: $webDeployPackage"
            $task.WebDeployPackage = $webDeployPackage
        }
    }
    
    # Convert PSObject with modified artifact filepath to json format
    $releaseDefJson = @($releaseDef) | ConvertTo-Json -Depth 99

    # try updating release definition with releaseDefJson, catch any errors in \\itd-pw46112\d$\Release Management\log\
    try { $updateReleaseDef = Invoke-RestMethod -Uri $Url -Method Put -Body $releaseDefJson -ContentType "application/json" -Headers $Headers }
    catch {
        $errorMsg = "PROJECT: $projectName , RELEASE ID: $releaseID `r`n" + $_ + "`r`n" 
        $errorMsg | Out-File -FilePath $UpdateReleaseDefLog -Append
    }


    # Write to console the updated artifact filepath in the json file that was used in the http PUT request
    if ($task.SourcePath.Length) { 
        $artifactFilePath = $updateReleaseDef.environments.deployPhases.workflowTasks.inputs.SourcePath
    }
    else {
        $artifactFilePath = $updateReleaseDef.environments.deployPhases.workflowTasks.inputs.WebDeployPackage
    }

    Write-Host "Artifact filepath has been updated to: " $artifactFilePath
    Write-Host "******************************************************************************* `n"
}

function main {
  # List of projects to update their pipelines, got project list json by using az devops CLI
  $projectList = Get-Content "D:\Release Management\Project Info\projectList.json" | ConvertFrom-Json

  # iterate through each project to update their build and release definitions
  foreach($project in $projectList.value.name) {
    Write-Host "Project: $project"

    $buildID = Get-BuildID $project
    Write-Host "Build ID(s): $buildID"

    $releaseID = Get-ReleaseID $project
    Write-Host "Release ID(s): $releaseID"
    # counter for ID elements in $releaseID array
    $rCount = 0
    
    if ($releaseID -eq $null) { 
        Write-Host "SKIPPING PROJECT. Project is empty or missing release pipeline. `n"
        "$project" | Out-File -FilePath "D:\Release Management\log\Skipped-Projects.txt" -Append
    }
    else {
        # if project has multiple builds (dev and prod), update both definitions
        foreach($build in $buildID) {
            Update-BuildDef $project $build
            Update-ReleaseDef $project $releaseID[$rCount++]
        }
    }
  }
} # end main function

main