# This script uses the Microsoft Graph and ExchangeOnlineManagement PowerShell modules to
# get the account information of all users from the office365 tenant with UserPrincipalName 
# ending in @domain.com or @domain2.com. 
#
# Filters the user list to find service accounts by searching for:
#       If the State field = "SERVICE ACCT"
#       If the DisplayName field matches specific string wildcard expressions defined in $ServiceAcctFilters
# Exports the two lists into one .xlsx excel workbook with 2 worksheets:
#       Worksheet 1: Service Accounts
#       Worksheet 2: ALL @domain.com and @domain2.com users


# Local filepaths
$MacPath = "/Users/user/Documents/OneDrive - Warner Bros. Entertainment Inc/scripts/output/InactiveUsers/"
$Certfp = "/Users/user/Desktop/PowerShellGraphCert.pfx"
$ReportOutputPath = "/Users/user/Documents/scripts/output/ServiceAccts.xlsx"
$MacKeychainKey = ""

# Get Service Principal cert file and encryption password from Mac Keychain app
$mypwd = security find-generic-password -s $MacKeychainKey -w
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($Certfp, $mypwd)

# Organization ID information for auth
$ClientId = ""
$TenantId = ""
$OrgName = "domain.onmicrosoft.com"

# Connect to your AzureAD tenant and Exchange Online
Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Certificate $cert
Select-MgProfile beta
Connect-ExchangeOnline -AppID $ClientId -Organization $OrgName -Certificate $cert

# Get all users that have UPN ending with @domain.com or @domain2.com
$UserDump = Get-MgUser -All -Property DisplayName,UserPrincipalName,GivenName,Surname,AccountEnabled,CreatedDateTime,AssignedLicenses,City,Country,UsageLocation,JobTitle,Department,MobilePhone,State,SignInActivity | `
            Where-Object { ($_.UserPrincipalName -like "*@domain.com" -or $_.UserPrincipalName -like "*@domain2.com") -and ($_.UserPrincipalName -notlike "adm-*") }

# Create new user list to filter information from $UserDump list
$NewUserList = @()
foreach ($user in $UserDump) {
    $UPN = $user.UserPrincipalName
    if($user.AssignedLicenses) {
        $MailboxType = (Get-EXOMailbox -Identity $UPN).RecipientTypeDetails
    } else {
        $MailboxType = "No Mailbox"
    }

    $NewUserList += [PSCustomObject]@{
        DisplayName = $user.DisplayName
        FirstName = $user.GivenName
        LastName = $user.Surname
        UserPrincipalName = $user.UserPrincipalName
        Country = $user.Country
        UsageLocation = $user.UsageLocation
        AccountEnabled = $user.AccountEnabled
        RecipientTypeDetails = $MailboxType
        CreatedDateTime = $user.CreatedDateTime
        LastSignInDateTime = $user.SignInActivity.LastSignInDateTime
        State = $user.State
        City = $user.City
        MobilePhone = [string]$user.MobilePhone
        JobTitle = $user.JobTitle
        Department = $user.Department
    }
}

# Filter $NewUserList to only get service accounts by searching for the defined filters
$ServiceAcct = @()
$ServiceAcctFilters = "US*", "UK*", "GBR*", "TV *", "Prod*", "FPP*", "*VFX*"
foreach ($filter in $ServiceAcctFilters) {
    $ServiceAcct += $NewUserList | Where-Object { ($_.DisplayName -like $filter) -or ($_.State -eq "SERVICE ACCT")}
}

# Remove any duplicate entries based on display name
$ServiceAcct_NoDupes = $ServiceAcct | Sort-Object -Unique -Property DisplayName | Select-Object *


# Export reports to excel format
$ServiceAcct_NoDupes | Export-Excel -Path $ReportOutputPath -WorksheetName "Service Accounts"
$NewUserList | Export-Excel -Path $ReportOutputPath -WorksheetName "ALL @domain and @domain2 users"
