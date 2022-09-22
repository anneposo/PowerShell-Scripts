# The script goes through csv list of users to block sign in and remove all assigned Microsoft licenses from their account.

# Local filepaths
$MacPath = "/Users/user/Documents/scripts/output/"
$OutputReport = $MacPath + "RemoveLicense-" + $(Get-Date -Format yyyy-MM-dd_HHmm) + ".csv"
$ErrorLog = $MacPath + "error/Error-RemoveLicense-" + $(Get-Date -Format yyyy-MM-dd_HHmm) + ".csv"
$LogPath = $MacPath + "log/RemoveLicense-" + $(Get-Date -Format yyyy-MM-dd_HHmm) + ".log"
$MacKeychainKey = ""

Start-Transcript -Path $LogPath -NoClobber

# Import CSV of the list of user to unlicense.
$userInfo = Import-Csv -Path ".\USER DUMP 4_6_2022.csv"

# Get Service Principal cert file and encryption password from Mac Keychain app
$mypwd = security find-generic-password -s $MacKeychainKey -w
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($Certfp, $mypwd)

# Connect to your AzureAD tenant and Exchange Online
Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Certificate $cert
Select-MgProfile beta
Connect-ExchangeOnline -AppID $ClientId -Organization $OrgName -Certificate $cert


# Function to block O365 sign in for input user
function Set-SignInBlocked {
    Param([string]$UPN)
    Try {
        Update-MgUser -UserId $UPN -AccountEnabled:$false
        Write-Output "Blocked sign in for $UPN"
    } Catch {
        $ErrorUsers += $user | Add-Member @{Status = $_.Exception.Message} -PassThru
        Write-Output "An error was logged for $UPN"
    }
}


$ErrorUsers = @() # Track users that output an error/cannot remove license
$UnlicensedUsers = foreach ($user in $userinfo) {
    $UPN = $user.UserPrincipalName

    $userList = Get-MgUser -UserId $UPN
    $LicensesToRemove = $userList.AssignedLicenses | Select -ExpandProperty SkuId

    # Block user sign in
    if ($userList.AccountEnabled) {
        Set-SignInBlocked $UPN
    }
    
    # Remove all Microsoft assigned licenses
    if ($LicensesToRemove) {
        Try {
            Set-MgUserLicense -UserId $UPN -RemoveLicenses $LicensesToRemove -AddLicenses @{}
            $user.AssignedLicenses = "Unlicensed"
            $user
        }Catch {
            $ErrorUsers += $user | Add-Member @{Status = $_.Exception.Message} -PassThru
        }
    }
}

# Export list of users that had license removed and skipped users
$UnlicensedUsers | Export-Csv -Path $OutputReport -NoTypeInformation
$ErrorUsers | Export-Csv -Path $ErrorLog -NoTypeInformation
