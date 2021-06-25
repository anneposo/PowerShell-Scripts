# Anne Poso 6/22/2021
#
# Gets list of all projects for given Azure DevOps organization in JSON format
# List is saved at .\output\projectList.json

# Requires VSTS PAT with scope User Profile and Project (READ)
$token = "Personal Access Token"
$authentication = [Text.Encoding]::ASCII.GetBytes(":$token")
$authentication = [System.Convert]::ToBase64String($authentication)
$headers = @{
    Authorization = ("Basic {0}" -f $authentication)
}

$organizationName = "Org-Name" # Name of Azure DevOps organization

function Get-ProjectList {
    # Change top value if project count is higher than 500
    $Url = "https://dev.azure.com/" + $organizationName + "/_apis/projects?`$top=500&api-version=6.0"
    $projectList = Invoke-RestMethod -Method Get -Uri $Url -Headers $headers
    $($projectList | ConvertTo-Json -Depth 99) | Out-File -FilePath ".\output\projectList.json"
}

Get-ProjectList