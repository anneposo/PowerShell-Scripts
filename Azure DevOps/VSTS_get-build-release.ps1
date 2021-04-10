# Anne Poso 10/15/2020
#
# NOTE: Removed any identifying organization and server names.
#
# Powershell script that uses Azure Devops Services REST API 5.1 to loop through a list of organization's
# Azure DevOps projects to get a project's current build and release definitions in json format.


# JSON definitions are saved in the file paths below
$buildDefPath = "D:\Release Management\Builds\"
$releaseDefPath = "D:\Release Management\Release\"

# File paths for script try-catch error logs
$getBuildDefError = "D:\Release Management\log\Get-BuildDef-error.txt"
$getReleaseDefError = "D:\Release Management\log\Get-ReleaseDef-error.txt"

$skippedBuild = "D:\Release Management\log\Empty-Build.txt"
$skippedRelease = "D:\Release Management\log\Empty-Release.txt"
$skippedProject = "D:\Release Management\log\Empty-Projects.txt"

# Requires VSTS PAT with build and release read access
$Token = Read-Host -Prompt "Please enter your VSTS PAT. (Must have build and release read access)" -AsSecureString
$bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Token)
$Token = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)


# VSTS PAT needs to be a base64 string to be used
$Authentication = [Text.Encoding]::ASCII.GetBytes(":$Token")
$Authentication = [System.Convert]::ToBase64String($Authentication)
$Headers = @{
    Authorization = ("Basic {0}" -f $Authentication)
}

# returns array of project's build IDs (compiled apps have 2 build IDs)
function Get-BuildID($projectName) {
    $Url = "https://organizationName.com/" + $projectName + "/_apis/build/definitions?api-version=5.1"

    #Get project's list of build definitions in PSObject format
    $buildList = Invoke-RestMethod -Method Get -Uri $Url -Headers $Headers

    # DEBUGGING: Output build list to json to see properties so that we can reference the lines we want to change/update
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

function Get-BuildDef($projectName, $buildID) {
    $Url = "https://organizationName.com/" + $projectName + "/_apis/build/definitions/" + $buildID + "?api-version=5.1"
    Write-Host "BUILD definition URL: $Url"

    # Get project build definition in PSObject format
    try { $buildDef = Invoke-RestMethod -Uri $Url -Method Get -Headers $Headers }
    catch { 
        $errorMsg = "PROJECT: $projectName , BUILD ID: $releaseID `r`n" + $_ + "`r`n"
        $errorMsg | Out-File -FilePath $getBuildDefError -Append
    }
    
    #Convert PSObject to json format
    $buildFilePath = $buildDefPath + $projectName + "-" + $buildID + "-BuildDef.json"
    $($buildDef | ConvertTo-Json -Depth 99) | Out-File -FilePath $buildFilePath
    Write-Host "Got definition: $projectName-BuildDef.json"
}

function Get-ReleaseDef($projectName, $releaseID) {
    $Url = "https://vsrm.dev.azure.com/organizationName/" + $projectName + "/_apis/release/definitions/" + $releaseID + "?api-version=5.1"
    Write-Host "RELEASE definition URL: $Url"  

    try { $releaseDef = Invoke-RestMethod -Uri $Url -Method Get -Headers $Headers }
    catch { 
        $errorMsg = "PROJECT: $projectName , RELEASE ID: $releaseID `r`n" + $_ + "`r`n"
        $errorMsg | Out-File -FilePath $getReleaseDefError -Append
    }

    $releaseFilePath = $releaseDefPath + $projectName + "-" + $releaseID + "-ReleaseDef.json"
    $($releaseDef | ConvertTo-Json -Depth 99) | Out-File -FilePath $releaseFilePath   
    Write-Host "Got definition: $projectName-ReleaseDef.json"
}

function main {
  # got project list json by using az devops CLI
  $projectList = Get-Content "D:\Release Management\Project Info\projectList.json" | ConvertFrom-Json

  # iterate through each project to get their build and release definitions
  foreach($project in $projectList.value.name) {
    Write-Host "Project: $project"

    $buildID = Get-BuildID $project
    Write-Host "Build ID(s): $buildID"

    $releaseID = Get-ReleaseID $project
    Write-Host "Release ID(s): $releaseID"
    # counter for ID elements in $releaseID array
    $rCount = 0

    # Projects missing pipelines will not have a json definition file, project names will be saved to the text files below (skippedprojects)
    if ($buildID -eq $null -AND $releaseID -eq $null) { "$project" | Out-File -FilePath $skippedProject -Append }
    elseif($buildID -eq $null) { "$project" | Out-File -FilePath $skippedBuild -Append }
    elseif($releaseID -eq $null){ "$project" | Out-File -FilePath $skippedRelease -Append }
    
    
    # if project has multiple builds (dev and prod), get both definitions
    foreach($build in $buildID) {
        Get-BuildDef $project $build
        Get-ReleaseDef $project $releaseID[$rCount++]
    }
    
  }

}

main