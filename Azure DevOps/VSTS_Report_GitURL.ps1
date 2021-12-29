# Generates report of an organization's Azure DevOps repository names and the repo URL

$Account ="AzureDevOpsAccountName"
$Accesstoken=" "
$passkey = ":$($Accesstoken)"

$encodedKey = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($passkey))
$token = "Basic $encodedKey"

$list = @()
$url = "https://$Account.visualstudio.com/_apis/git/repositories?api-version=5.1"
$result = Invoke-RestMethod $url -Method Get -Headers @{ Authorization = $token }
    
if ($result.count -gt 0){
    $result.value | % {
        $list += [pscustomobject]@{
            Project = $_.project.name
            url = $_.remoteUrl
        }
    }
}

$list| Sort-Object -Property Project |Export-Csv -Path "C:\FilePath\giturl.csv" -NoTypeInformation