<#
    .SYNOPSIS
		Gets a list of inactive @domain.com and @domain2.com users from organization's AzureAD 
		tenant whose last interactive sign-in was over 6 months ago.
	
	.DESCRIPTION
		Gets a list of inactive @domain.com and @domain2.com users from organziation's AzureAD 
		tenant whose last interactive sign-in was over 6 months ago. The inactive user list is 
		separated into different Excel worksheets according to the team the user belongs to (US, 
		UK, or Accounting).

		The user lists are exported to Excel workbook format and uploaded to a filepath on the
        local machine specified by $FilepathAll, $FilepathUSPT, $FilepathUKPT, and $FilepathAcct.

    .NOTES
        AUTHOR: Anne Poso
        LASTEDIT: June 15, 2022
#>

# Local filepaths
$MacPath = "/Users/user/Documents/scripts/output/InactiveUsers/"
$Certfp = '/Users/user/Desktop/PowerShellGraphCert.pfx'

$LogPath = $MacPath + "log/InactiveUsers " + $(Get-Date -Format yyyy-MM-dd_HHmm) + ".log"
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

# Get date for checking how long user is inactive
$Today = (Get-Date)
$April2020 = Get-Date -Date "04/01/2020" # For LastSignInDateTime property


# Get list of all users in tenant
Write-Output "Getting all users in domain.com and domain2.com domains..."
$UserDump = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,GivenName,Surname,AccountEnabled,CreatedDateTime,AssignedLicenses,City,Country,UsageLocation,JobTitle,Department,MobilePhone,State,SignInActivity | `
               Where-Object { ($_.UserPrincipalName -like "*@domain.com" -or $_.UserPrincipalName -like "*@domain2.com") -and ($_.UserPrincipalName -notlike "adm-*") }
$MailboxDump = Get-EXOMailbox -ResultSize unlimited
Write-Output "Finished getting all users"


# Get list of user IDs part of specific o365 security groups
$ServiceAcctList = (Get-MgGroupMember -GroupId $((Get-MgGroup -Filter "DisplayName eq 'Service Accounts'").Id)).Id
$ITGroupList = (Get-MgGroupMember -GroupId $((Get-MgGroup -Filter "DisplayName eq 'IT Engineers'").Id)).Id
$AcctGroups = "ACCT 1", "ACCT 2", "ACCT 3", "ACCT 4"
$AccountingGroupList = foreach ($group in $AcctGroups) {
    (Get-MgGroupMember -All -GroupId $((Get-MgGroup -Filter "DisplayName eq '$group'").Id)).Id
}


# Get list of inactive users from the user list
Write-Output "Starting process to get inactive users..."
$InactiveUserList = foreach ($user in $UserDump) {
    $UPN = $user.UserPrincipalName
    $isInactive = $false

    # If user is not a service account or IT engineer, then check for inactivity
    if ($user.Id -notin $ServiceAcctList -and $user.Id -notin $ITGroupList) {
        # Filter for only inactive users
        if($user.SignInActivity.LastSignInDateTime) {
            $LastLogin = $user.SignInActivity.LastSignInDateTime
            
            # Create a timespan to calculate the number of days an user has not signed in interactively
            $TimeSpan = New-TimeSpan -Start $LastLogin -End $Today
            $DaysInactive = $TimeSpan.Days

            if ($DaysInactive -gt 180) { # If user has not signed in in the last 6 months
                $isInactive = $true
            }

        } else { # Blank SignInActivity means the user never logged in or last login was before April 2020
            if ($user.CreatedDateTime -gt $April2020) { # If user was created after April 2020, then they never logged in
                $LastLogin = $user.CreatedDateTime
                $TimeSpan = New-TimeSpan -Start $LastLogin -End $Today
                $DaysInactive = $TimeSpan.Days

                if ($DaysInactive -gt 180) {
                    $isInactive = $true
                }

            } else { # If user was created before April 2020
                $LastLogin = "Last login was before April 2020"
                $TimeSpan = New-TimeSpan -Start $April2020 -End $Today
                $DaysInactive = [string]$TimeSpan.Days + '+'
                $isInactive = $true
            }
            
        }
    }

    # If $isInactive flag is true, add user to $InactiveUserList
    if ($isInactive) {
        # Convert SkuId from AssignedLicenses property to actual license name
        $Licenses=@()
        if($user.AssignedLicenses) {
            $SkuId = $user.AssignedLicenses.SkuId
            foreach ($license in $SkuId) {
                switch($license) {
                    "05e9a617-0261-4cee-bb44-138d3ef5d965" { $Licenses += "Microsoft 365 E3"}
                    "66b55226-6b4f-492c-910c-a3b7a3c9d993" { $Licenses += "Microsoft 365 F3"}
                    "18181a46-0d4e-45cd-891e-60aabd171b4e" { $Licenses += "Office 365 E1"}
                    "710779e8-3d4a-4c88-adb9-386c958d1fdf" { $Licenses += "Microsoft Teams Exploratory"}
                    "f30db892-07e9-47e9-837c-80727f46fd3d" { $Licenses += "Microsoft Power Automate Free"}
                    "1f2f344a-700d-42c9-9427-5cea1d5d7ba6" { $Licenses += "Microsoft Stream Trial"}
                    default { $Licenses += $license }
                }
            }
            $Licenses = $Licenses -join "+"
        } else { $Licenses = "Unlicensed" }

        # Not using search for user's mailbox info below because Where-Object will slow down script runtime a lot
        <# if ($UPN -in $MailboxDump.UserPrincipalName) {
            $Mailbox = $($MailboxDump | Where-Object { $_.UserPrincipalName -eq $UPN }).RecipientTypeDetails
        } else { $Mailbox = "None" } #>

        # If user has a mailbox, get the mailbox type (User or Shared) - using foreach than where-object for faster search
        $Mailbox = $null
        foreach ($box in $MailboxDump) {
            if ($box.UserPrincipalName -eq $UPN) {
                $Mailbox = $box.RecipientTypeDetails
                continue
            }
        }
        if (!$Mailbox) { $Mailbox = "None"}

        if($Mailbox -eq "UserMailbox") { # Only include regular user mailboxes in report, excludes shared mailboxes and no mailboxes
            #Write-Output "$UPN to be disabled true`r`nLast logon: $LastLogin $DaysInactive days ago`r`n"
            [PSCustomObject]@{
                DisplayName = $user.DisplayName
                FirstName = $user.GivenName
                LastName = $user.Surname
                UserPrincipalName = $user.UserPrincipalName
                Country = $user.Country
                UsageLocation = $user.UsageLocation
                AccountEnabled = $user.AccountEnabled
                AssignedLicenses = $Licenses
                MailboxType = $Mailbox
                CreatedDateTime = [string]$user.CreatedDateTime
                LastSignInDateTime = [string]$LastLogin
                DaysInactive = [string]$DaysInactive
                ProductionName = $user.State
                PersonalEmail = $user.City
                MobilePhone = [string]$user.MobilePhone
                JobTitle = $user.JobTitle
                Department = $user.Department
                ObjectId = $user.Id
            }
        }
        
    }
}

# Split inactive users into different worksheets based on team the user belongs to - US, UK, Accounting
$USList = @()
$UKList = @()
$AccountingList = @()
$OtherList = @()


Write-Output "Sorting users into US, UK, and Accounting lists..."
foreach ($user in $InactiveUserList) {
    $isOther = $false
    if(($user.Department -like "*Acc*") -or ($user.JobTitle -like "*Accountant*") -or ($user.JobTitle -like "*Accounting*") -or ($user.ObjectId -in $AccountingGroupList)) {
        $AccountingList += $user
    }
    elseif($user.Country) {
        if($user.Country -like "*States*" -or $user.Country -like "US*" -or $user.Country -like "AU*") {
            $USList += $user
        } elseif ($user.Country -like "UK*" -or $user.Country -like "United K*") {
            $UKList += $user
        } else {
            $isOther = $true
        }
    } else {
        $isOther = $true
    }

    # If user has empty Country field, look at Usage Location field
    if($isOther) {
        if($user.UsageLocation -eq "US" -or $user.UsageLocation -eq "AU") {
            $USList += $user
        } elseif ($user.UsageLocation -eq "GB" -or $user.UsageLocation -eq "DE"){ # DE = Germany
            $UKList += $user
        } else {
            $OtherList += $user
        }
    }
}

# Uses PST timezone for report's filename
$pstzone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Pacific Standard Time")
$psttime = [System.TimeZoneInfo]::ConvertTimeFromUtc((Get-Date).ToUniversalTime(), $pstzone)
$ReportTimestamp = Get-Date $psttime -Format yyyy-MM-dd_HHmm


# Export user lists from PSObject to Excel document with multiple worksheets
Write-Output "Exporting user lists..."

# File with all worksheets in one workbook
$FilepathAll = $MacPath + "InactiveUsers " + $ReportTimestamp + ".xlsx"
Try {
    $InactiveUserList | Export-Excel -Path $FilepathAll -WorksheetName "ALL Inactive Users"
    $USList | Export-Excel -Path $FilepathAll -WorksheetName "US Inactive"
    $UKList | Export-Excel -Path $FilepathAll -WorksheetName "UK Inactive"
    $AccountingList | Export-Excel -Path $FilepathAll -WorksheetName "Accounting Inactive"
    $OtherList | Export-Excel -Path $FilepathAll -WorksheetName "Unknown Inactive"
} Catch {
    Write-Error -Message $_.Exception.Message
	Write-Output "Error exporting user lists to $FilepathAll."
}

# File to send to US team
$FilepathUSPT = $MacPath + "US InactiveUsers " + $ReportTimestamp + ".xlsx"
Try {
	$USList | Export-Excel -Path $FilepathUSPT -WorksheetName "US Inactive"
} Catch {
	Write-Error -Message $_.Exception.Message
	Write-Output "Error exporting user lists to $FilepathUSPT."
}

# File to send to UK team
$FilepathUKPT = $MacPath + "UK InactiveUsers " + $ReportTimestamp + ".xlsx"
Try {
	$UKList | Export-Excel -Path $FilepathUKPT -WorksheetName "UK Inactive"
} Catch {
	Write-Error -Message $_.Exception.Message
	Write-Output "Error exporting user lists to $FilepathUKPT."
}

# File to send to Accounting team
$FilepathAcct = $MacPath + "Acct InactiveUsers " + $ReportTimestamp + ".xlsx"
Try {
	$AccountingList | Export-Excel -Path $FilepathAcct -WorksheetName "Accounting Inactive"
} Catch {
	Write-Error -Message $_.Exception.Message
	Write-Output "Error exporting user lists to $FilepathAcct."
}

<#
# Upload exported Excel document to Azure blob storage
$Filelist = $FilepathAll, $FilepathUSPT, $FilepathUKPT, $FilepathAcct
foreach ($file in $Filelist) {
	Write-Output "Uploading $file to storage account in o365-reports container..."
	$Blobreport = @{
		File = $file
		Container = 'o365-reports'
		Blob = $file
		Context = $StorageAccount
		StandardBlobTier = 'Cool'
	}

	Try {
		Set-AzStorageBlobContent @Blobreport -Force
		Write-Output "$file successfully uploaded."
	} Catch {
		Write-Error -Message $_.Exception.Message
		Write-Output "Error uploading $file."
	}
}
#>

Disconnect-MgGraph
Disconnect-ExchangeOnline -Confirm:$false
Stop-Transcript