<#
    .SYNOPSIS
	Imports an excel list of users and goes through each user to assign an E3 Microsoft
        license, enable litigation hold, and finally delete the O365 user account.
	
	.DESCRIPTION
        The script imports an excel list of users and first checks if there are at least 75 E3 
        Microsoft licenses available to assign. If there are enough, it will work in batches of 
        50 users at a time to assign the E3 license, enable litigation hold, then delete the user 
        to free up the license again for the next batch.

	Users are deleted with a 6 month litigation hold and their mailbox is converted to an
        inactive mailbox to hold their associated mailbox data until the litigation hold expires.

        The script will export two reports:
            DeletedUsers-yyyy-MM-dd.csv             - List of users that were successfully deleted
            ErrorLog-DeletedUsers-yyyy-MM-dd.csv    - List of users that output an error on action calls
#>

# Local filepaths
$MacPath = "/Users/user/Documents/scripts/cleanup/"
$DeletedUsersReport = $MacPath + "DeletedUsers-" + $(Get-Date -Format yyyy-MM-dd_HHmm) + ".csv"
$DeletedUsersErrorLog = $MacPath + "ErrorLog-DeletedUsers-" + $(Get-Date -Format yyyy-MM-dd_HHmm) + ".csv"
$Certfp = '/Users/user/Desktop/PowerShellGraphCert.pfx'
$MacKeychainKey = ""

$LogPath = $MacPath + "log/RemoveUsers-" + $(Get-Date -Format yyyy-MM-dd_HHmm) + ".log"
Start-Transcript -Path $LogPath -NoClobber

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

# Import user list from excel file
$UserReportPath = $MacPath + "test.xlsx"
$UsersToRemove = Import-Excel -Path $UserReportPath

# Function to write output with current timestamp
function Write-LogOutput {
    param([string]$Text)
    Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss tt") : $Text"
}

# Check if there are enough E3 Licenses available to run this script. Need at least 50, using buffer of 25, so at least 75 must be availble
$e3LicenseDetail = Get-MgSubscribedSku | Where-Object {$_.SkuPartNumber -eq "SPE_E3"} | Select SkuPartNumber,ConsumedUnits,PrePaidUnits
$e3LicenseCount = $e3LicenseDetail.PrepaidUnits.Enabled - $e3LicenseDetail.ConsumedUnits
Write-LogOutput -Text "There are $e3LicenseCount E3 licenses available."
if ($e3LicenseCount -lt 75) {
    Write-LogOutput -Text "There are less than 75 E3 licenses available. Please free up additional E3 licenses before running this script."
    exit
}

# Start process to remove users
$ErrorUsers = @()
$TotalDeletedUsers = @() # Track total number of users that are deleted
$BatchSize = 50 # Size of batch to process a number of users at a time

# For loop to iterate through all users that need to be deleted
for ($i = 0; $i -lt $UsersToRemove.count; $i+=$BatchSize) {
    # Create sublist batch of X users to remove at a time from complete user list
    $BatchToRemove = $UsersToRemove[$i .. ($i+($BatchSize - 1))]

    # Iterate through batch first time to assign e3 license and enable lit hold
    foreach ($user in $BatchToRemove) {
        $UPN = $user.UserPrincipalName
        
        # First check if user is assigned an E3 License, if not, assign one
        $UserLicense = (Get-MgUser -UserID $UPN).AssignedLicenses
        if(!$UserLicense) {
            Write-LogOutput -Text "$UPN has no licenses assigned."
            Set-MgUserLicense -UserId $UPN -AddLicenses @{SkuId = "05e9a617-0261-4cee-bb44-138d3ef5d965"} -RemoveLicenses @()
            Write-LogOutput -Text "Completed assigning E3 License for $UPN."
        } elseif ($UserLicense.SkuId -eq "05e9a617-0261-4cee-bb44-138d3ef5d965") {
            Write-LogOutput -Text "$UPN is already assigned an E3 license."
        } else { # If user has assigned license, but E3 not found, unassign all licenses and assign E3 license
            $LicensesToRemove = $UserLicense.SkuId
            Set-MgUserLicense -UserId $UPN -RemoveLicenses $LicensesToRemove -AddLicenses @{SkuId = "05e9a617-0261-4cee-bb44-138d3ef5d965"}
            Write-LogOutput -Text "Completed assigning E3 License for $UPN."
        }
    }

    Write-LogOutput -Text "Sleeping for 5 minutes to wait for license changes to update.."
    Start-Sleep -Seconds 300

    foreach ($user in $BatchToRemove) {
        $UPN = $user.UserPrincipalName
        # Wait for mailbox to be provisioned after assigning license
        do {
            $mailbox = $Null
            Write-LogOutput -Text "Checking if $UPN mailbox is provisioned..."
            Start-Sleep -seconds 5
            $mailbox = (Get-MgUser -Userid $UPN -Property ProvisionedPlans).ProvisionedPlans
        } while(!$mailbox)
        Write-LogOutput -Text "$UPN is provisioned."

        # Once mailbox is provisioned, enable litigation hold with 180 day duration
        Try {
            Set-Mailbox $UPN -LitigationHoldEnabled $true -LitigationHoldDuration 180 -ErrorAction Stop
            Write-LogOutput -Text "Enabled litigation hold for $UPN."
        } Catch {
            $ErrorUsers += $user | Add-Member @{LitHoldStatus = $_.Exception.Message} -PassThru
            Write-LogOutput -Text "Failed to enable litigation hold for $UPN. Check error log for more details.."
        }
    }

    # Iterate through batch second time to wait until lit hold is enabled
    $LitHoldEnabled = $false
    while(!$LitHoldEnabled) {
        Write-LogOutput -Text "Checking if all users in the batch have litigation hold enabled..."
        foreach ($user in $BatchToRemove) {
            $UPN = $user.UserPrincipalName
            if ($((Get-Mailbox -Identity $UPN).LitigationHoldEnabled)){
                $LitHoldEnabled = $true
                Write-LogOutput -Text "$UPN has litigation hold"
            } else {
                $LitHoldEnabled = $false
                Write-LogOutput -Text "$UPN does NOT have litigation hold enabled, waiting 5 minutes before checking again..."
                Start-Sleep -Seconds 5 # Wait 5 minutes to check again
                break
            }
        }
    }
    # If while loop exits, then all users in the batch have lit hold enabled
    Write-LogOutput -Text "All users in the batch have litigation hold enabled, proceeding with removing user..."

    # Proceed with deleting the batch of users
    $BatchDeletedUsers = [System.Collections.Generic.List[Object]]::new()
    foreach ($user in $BatchToRemove) {
        $UPN = $user.UserPrincipalName
        Try {
            Remove-MgUser -UserId $UPN -ErrorAction Stop
            $user | Add-Member @{Status = "Deleted"; LitigationHoldEnabled = "True"; DateDeleted = $(Get-Date); LitigationHoldExpires = $((Get-Date).AddDays(180))} -PassThru
            $BatchDeletedUsers.Add($user)
            Write-LogOutput -Text "Deleted user: $UPN."
        } Catch {
            $ErrorUsers += $user | Add-Member @{Status = $_.Exception.Message} -PassThru
            Write-LogOutput -Text "Failed to delete user $UPN. Check error log for more details.."
        }
    }

    # Append batch of users to deleted users report
    #$BatchDeletedUsers | Export-Csv -Path $DeletedUsersReport -Append
    $TotalDeletedUsers += $BatchDeletedUsers

    # Reached end of batch, move on to next batch of users in for loop
}

$TotalDeletedUsers | Export-Csv -Path $DeletedUsersReport -NoTypeInformation
$ErrorUsers | Export-Csv -Path $DeletedUsersErrorLog -NoTypeInformation

Stop-Transcript
