# Generates report of organization's repository user access list in Azure DevOps 

$Account ="AzureDevOpsAccountName"
$Accesstoken=" "

$passkey = ":$($Accesstoken)"

$encodedKey = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($passkey))
$token = "Basic $encodedKey"

$projectsResult = Invoke-RestMethod "https://$Account.visualstudio.com/_apis/projects?`$top=500&api-version=5.1" -Method Get -Headers @{ Authorization = $token }

$list = @()

$projectsResult.value | % {
    $project = $_.name
    $projectID = $_.id

    $url = "https://$Account.visualstudio.com/_apis/projects/$projectID/teams/$project Team/members?api-version=5.1"

    $result = Invoke-RestMethod $url -Method Get -Headers @{ Authorization = $token }

    if ($result.count -gt 0){
        $result.value | % {
            $list += [pscustomobject]@{
                Project = $project
                member=$_.identity.displayName
                email=$_.identity.uniqueName
            }
        }
    }
}

$list| Sort-Object -Property Project|Export-Csv -Path "C:\FilePath\userlist.csv" -NoTypeInformation