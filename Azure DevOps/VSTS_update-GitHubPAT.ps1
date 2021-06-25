# Anne Poso 6/22/2021
# 
# PowerShell script that uses Azure DevOps Services REST API 6.0 to loop through a list of
# VSTS projects and update the GitHub service connection with new Personal Access Token (PAT).
#
# NOTE: Update variables $token, $organizationName, $githubToken with your appropriate values.

# Requires VSTS PAT with Service Connections scope (READ, QUERY, & MANAGE)
$token = "Azure DevOps Personal Access Token"
$authentication = [Text.Encoding]::ASCII.GetBytes(":$token")
$authentication = [System.Convert]::ToBase64String($authentication)
$headers = @{
    Authorization = ("Basic {0}" -f $authentication)
}

$organizationName = "Org-Name" # Name of Azure DevOps organization
$githubToken = "GitHub Personal Access Token" # GitHub PAT to authorize Azure DevOps GitHub service connection
$updateConnLog = ".\output\UpdateServiceConnection-Error.txt" # Path to output error logs

# Get project's existing service connections and return GitHub service connection ID
function Get-ServiceConnection($projectName) {
    $Url = "https://dev.azure.com/" + $organizationName + "/" + $projectName + "/_apis/serviceendpoint/endpoints?api-version=6.0-preview.4"
    $connList = Invoke-RestMethod -Method Get -Uri $Url -Headers $Headers
    #$($connList | ConvertTo-Json -Depth 99) | Out-File -FilePath ".\output\testConn.json"

    # Find and return GitHub service connection ID
    $githubConn = $connList.value | where {$_.type -eq "github"} | select id, name
    return $githubConn
}

# Update GitHub service connection with new PAT
function Update-ServiceConnection($projectName, $connID) {
    # Get JSON object of existing GH service connection for request body
    $githubConnUrl = "https://dev.azure.com/" + $organizationName + "/" + $projectName + "/_apis/serviceendpoint/endpoints?endpointNames=" + $connID.name + "&api-version=6.0-preview.4"
    $githubConnInfo = Invoke-RestMethod -Method Get -Uri $githubConnUrl -Headers $Headers

    # Path to save GitHub service connection JSON definition before changes are made
    $githubConnInfoPath = ".\output\github_SCInfo_backup\" + $projectName + "-GHConnInfo.json"
    $($githubConnInfo | ConvertTo-Json -Depth 99) | Out-File -FilePath $githubConnInfoPath

    # New GitHub PAT to update on VSTS service connection
    $tokenParam = @{
        accessToken = $githubToken
    }

    $githubConnInfo.value.authorization.scheme = "PersonalAccessToken"
    $githubConnInfo.value.authorization | Add-Member -MemberType NoteProperty -Name "parameters" -Value $tokenParam -Force -PassThru

    # Convert from PS object to JSON for request body
    $connInfoJson = $githubConnInfo.value | ConvertTo-Json -Depth 99

    # Send PUT request to update the project's GitHub service connection
    $Url = "https://dev.azure.com/" + $organizationName + "/" + $projectName + "/_apis/serviceendpoint/endpoints/" + $connID.id + "?api-version=6.0-preview.4"
    try { 
        Invoke-RestMethod -Uri $Url -Method Put -Body $connInfoJson -ContentType "application/json" -Headers $Headers 
    } catch {
        $errorMsg = "PROJECT: $projectName , SERVICE CONNECTION: $connID `r`n" + $_ + "`r`n"
        $errorMsg | Out-File -FilePath $updateConnLog -Append
        Write-Host "An error was logged for" $projectName "in" $updateConnLog
    }
    if (!$errorMsg) { Write-Host "Successfully updated GitHub service connection with new PAT."}
}

function main {
    # Ran VSTS_Get-ProjectList.ps1 to get latest project list
    $projectList = Get-Content ".\output\projectList.json" | ConvertFrom-Json

    foreach($projectName in $projectList.value.name) {
        Write-Host "Checking project:" $projectName
        # Get existing GitHub service connection
        $githubConn = Get-ServiceConnection $projectName

        # Update GitHub service connection if one exists, else skip it
        if ($null -ne $githubConn) {
            Update-ServiceConnection $projectName $githubConn
        } else { 
            Write-Host "Project" $projectName "has no GitHub service connection."
        }
    }
}

main