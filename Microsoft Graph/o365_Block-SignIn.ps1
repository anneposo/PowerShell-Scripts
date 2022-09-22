# The script goes through excel list of users and blocks their sign in.

# Local filepaths
$MacPath = "/Users/user/Documents/scripts/cleanup/deactivated/"
$ReportFilepath = $MacPath + "test.xlsx"
$DeactivatedUsersReport = $MacPath + "Deactivated Users " + $(Get-Date -Format yyyy-MM-dd_HHmm) + ".xlsx"
$ErrorLogFilepath = $MacPath + "ErrorLog-BlockSignIn-" + $(Get-Date -Format yyyy-MM-dd_HHmm) + ".xlsx"
$Certfp = '/Users/user/Desktop/PowerShellGraphCert.pfx'
$MacKeychainKey = ""

# Start log transcript
$LogPath = $MacPath + "log/Deactivated Users " + $(Get-Date -Format yyyy-MM-dd_HHmm) + ".log"
Start-Transcript -Path $LogPath -NoClobber


# Get Service Principal cert file and encryption password from Mac Keychain app
$mypwd = security find-generic-password -s $MacKeychainKey -w
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($Certfp, $mypwd)


# Organization ID information for auth
$ClientId = ""
$TenantId = ""

# Connect to AzureAD tenant
Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Certificate $cert

# Import list of users to deactivate
$UsersToDeactivate = Import-Excel -Path $ReportFilepath

# Initialize variables to track list of users
$ErrorUsers = [System.Collections.Generic.List[Object]]::new()
$DeactivatedUsers = [System.Collections.Generic.List[Object]]::new()

foreach ($user in $UsersToDeactivate) {
    $UPN = $user.UserPrincipalName
    $IsEnabled = (Get-MgUser -UserId $UPN -Property AccountEnabled).AccountEnabled

    if($IsEnabled) {
        try {
            Update-MgUser -UserId $UPN -AccountEnabled:$false -ErrorAction Stop
            Write-Output "Deactivated $UPN."
            $DeactivatedUsers.Add($user)
        } catch {
            Write-Output "Failed to deactivate $UPN. Check error log for more details."
            $ErrorMsg = [PSCustomObject]@{
                UserPrincipalName = $UPN
                Error = $_.Exception.Message
            }
            $ErrorUsers.Add($ErrorMsg)
        }
    } else {
        Write-Output "$UPN is already deactivated."
        $DeactivatedUsers.Add($user)
    }
    
}

$DeactivatedUsers | Export-Excel -Path $DeactivatedUsersReport
$ErrorUsers | Export-Excel -Path $ErrorLogFilepath

Disconnect-MgGraph
Stop-Transcript