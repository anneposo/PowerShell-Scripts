# Anne Poso 10/15/2020
#
# NOTE: Removed any identifying organization and server names.
#
# Requires user to have backup JSON pipeline definitions to revert back to
# (usually run VSTS_get-build-release.ps1 *BEFORE* running update script to get backup JSON definitions)
# Reverts any changes made to VSTS projects before executing VSTS_update-build-release.ps1
# Uses directories D:\Release Management\Build\ and D:\Release Management\Release\ for JSON definitions


# File paths for script try-catch error logs
$UpdateBuildDefLog = "D:\Release Management\log\Revert-BuildDef-error.txt"
$UpdateReleaseDefLog = "D:\Release Management\log\Revert-ReleaseDef-error.txt"

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
    $buildList = Invoke-RestMethod -Method Get -Uri $Url -Headers $Headers
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

    $buildFilePath = "D:\Release Management\Builds\" + $projectName + "-" + $buildID + "-BuildDef.json"
    $buildDef = Get-Content $buildFilePath

    # try updating build definition with json, catch any errors
    try { $updateBuildDef = Invoke-RestMethod -Uri $Url -Method Put -Body $buildDef -ContentType "application/json" -Headers $Headers }
    catch {
        $errorMsg = "PROJECT: $projectName , BUILD ID: $buildID `r`n" + $_ + "`r`n"
        $errorMsg | Out-File -FilePath $UpdateBuildDefLog -Append
    }
    
    Write-Host "Project build definition reverted to: " $buildFilePath
}


# Updates release definition's artifact source filepath value
function Update-ReleaseDef($projectName, $releaseID) {
    Write-Host "=============================UPDATING RELEASE ID $releaseID============================="

    $Url = "https://vsrm.dev.azure.com/organizationName/" + $projectName + "/_apis/release/definitions/" + $releaseID + "?api-version=5.1"
    Write-Host "Update RELEASE definition URL: $Url"  
    
    $releaseFilePath = "D:\Release Management\Release\" + $projectName + "-" + $releaseID + "-ReleaseDef.json"
    $releaseDef = Get-Content $releaseFilePath

    # try updating release definition with releaseDefJson, catch any errors in \\itd-pw46112\d$\Release Management\log\
    try { $updateReleaseDef = Invoke-RestMethod -Uri $Url -Method Put -Body $releaseDef -ContentType "application/json" -Headers $Headers }
    catch {
        $errorMsg = "PROJECT: $projectName , RELEASE ID: $releaseID `r`n" + $_ + "`r`n"
        $errorMsg | Out-File -FilePath $UpdateReleaseDefLog -Append
    }

    Write-Host "Project release definition reverted to: " $releaseFilePath
    Write-Host "******************************************************************************* `n"
}

function main {
    # got project list json by using az devops CLI
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