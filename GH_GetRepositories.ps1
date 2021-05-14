# Gets list of all repositories in a GitHub private organization.
# Report includes repository name, URL, description, and collaborators.

Import-Module PowerShellForGitHub

$token = "Enter PAT here"

# Gets list of Github repositories in organization
$repoList = Get-GitHubRepository -OrganizationName "OrgName" -Type All -AccessToken $token | select name, RepositoryUrl, description

# Gets list of users with access to repository
$count = 0
foreach ($repo in $repoList.name) {
    $repoContributors = Get-GitHubRepositoryCollaborator -Ownername "OwnerName" -RepositoryName $repo -AccessToken $token | select username
    if ($count -eq 0) {
        $repoList | Add-Member -MemberType NoteProperty -Name 'Permissions' -Value $repoContributors
        $count++
    }
    else {
        $repoList[$count++].Permissions = $repoContributors
    }
}

$repoList | select name, RepositoryUrl, description, @{Name='Repository User Access List';Expression={$_.Permissions.username -join ';'}} | Export-Csv -Path "C:\Users\rposo\Desktop\GithubRepos.csv" -NoTypeInformation