<#
    .SYNOPSIS
		Powershell script that is hosted in Azure as a runbook in an Azure Automation Account.
		Gets a list of inactive @domain.com and @domain2.com users from organization's AzureAD 
		tenant whose last interactive sign-in was over 6 months ago.
	
	.DESCRIPTION
		Gets a list of inactive @domain.com and @domain2.com users from organziation's AzureAD 
		tenant whose last interactive sign-in was over 6 months ago. The inactive user list is 
		separated into different Excel worksheets according to the team the user belongs to (US, 
		UK, or Accounting).

		The user lists are exported to Excel workbook format and uploaded to an Azure storage 
		account in the o365-reports container.

		This runbook authenticates from the organization1 Azure tenant to organization2 AzureAD 
		tenant using a service principal that is hosted in the organization2 AzureAD tenant. 
		The certificate used for authentication will expire on 05/25/2024.

    .NOTES
        AUTHOR: Anne Poso
        LASTEDIT: June 15, 2022
#>

Connect-AzAccount -Identity | Out-null

# Get variables from Azure Automation Account for authentication to organization2 tenant
$Cert = Get-AzAutomationCertificate -ResourceGroupName "resource-group" -AutomationAccountName "AutomationAccount" -Name "Service Principal Name"
$ClientId = (Get-AzAutomationVariable -ResourceGroupName "resource-group" -AutomationAccountName "AutomationAccount" -Name sp_clientid).Value
$TenantId = (Get-AzAutomationVariable -ResourceGroupName "resource-group" -AutomationAccountName "AutomationAccount" -Name sp_tenantid).Value
$Organization = (Get-AzAutomationVariable -ResourceGroupName "resource-group" -AutomationAccountName "AutomationAccount" -Name Organization).Value

# Create Azure storage context to upload blobs to Azure storage account
$StorageAccount = New-AzStorageContext -StorageAccountName "storageaccount" -StorageAccountKey $(Get-AutomationVariable -Name office365_sa_key)

# Connect to organization2 AzureAD tenant using service principal authentication
Try {
	Write-Output "Connecting to MS Graph and Exchange..."
	Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $Cert.Thumbprint
	Select-MgProfile beta # SignInActivity attribute is not exposed in v1.0 API yet, so use beta
	Connect-ExchangeOnline -AppId $ClientId -Organization $Organization -CertificateThumbprint $Cert.Thumbprint
} Catch {
	Write-Error -Message $_.Exception.Message
    Break
}
Write-Output "Connect process done."

# Get date for checking how long user is inactive
$Today = (Get-Date)
$April2020 = Get-Date -Date "04/01/2020" # For LastSignInDateTime property

# Get list of all users in tenant
Write-Output "Getting all users in domain.com and domain2.com domains..."
$UserDump = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,GivenName,Surname,AccountEnabled,CreatedDateTime,AssignedLicenses,City,Country,UsageLocation,JobTitle,Department,MobilePhone,State,SignInActivity | `
               Where-Object { ($_.UserPrincipalName -like "*@domain.com" -or $_.UserPrincipalName -like "*@domain2.com") -and ($_.UserPrincipalName -notlike "adm-*") }
$MailboxDump = Get-EXOMailbox -ResultSize unlimited
Write-Output "Finished getting all users"

# Get list of service accounts using 'Service Accounts' o365 security group
$ServiceAcctList = (Get-MgGroupMember -GroupId $((Get-MgGroup -Filter "DisplayName eq 'Service Accounts'").Id)).Id


# Get list of inactive users from the user list
Write-Output "Starting process to get inactive users..."
$InactiveUserList = foreach ($user in $UserDump) {
    $UPN = $user.UserPrincipalName
    $isInactive = $false

    # If user is not a service account, then check for inactivity
    if ($user.Id -notin $ServiceAcctList) {
        # Filter for only inactive users by LastSignInDateTime property
        if($user.SignInActivity.LastSignInDateTime) {
            $LastLogin = $user.SignInActivity.LastSignInDateTime
            
			# Create a timespan to calculate the number of days an user has not signed in interactively
            $TimeSpan = New-TimeSpan -Start $LastLogin -End $Today
            $DaysInactive = $TimeSpan.Days

            if ($DaysInactive -gt 180) { # If user has not signed in in the last 6 months, set $isInactive flag
                $isInactive = $true
            }

        } else { # Blank SignInActivity means the user never logged in or last login was before April 2020
            if ($user.CreatedDateTime -gt $April2020) { # If user was created after April 2020, then they never logged in
                $LastLogin = $user.CreatedDateTime
                $TimeSpan = New-TimeSpan -Start $LastLogin -End $Today
                $DaysInactive = $TimeSpan.Days

                if ($DaysInactive -gt 180) { # If user has not signed in in the last 6 months, set $isInactive flag
                    $isInactive = $true
                }

            } else { # If user was created before April 2020 and LastSignInDateTime is blank, set $isInactive flag
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

		# If user has a mailbox, get the mailbox type (User or Shared) - using foreach than where-object for faster search
        $Mailbox = $null
        foreach ($box in $MailboxDump) {
            if ($box.UserPrincipalName -eq $UPN) {
                $Mailbox = $box.RecipientTypeDetails
                continue
            }
        }
        if (!$Mailbox) { $Mailbox = "None"}

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

# Split inactive users into different worksheets based on team the user belongs to - US, UK, Accounting
$USList = @()
$UKList = @()
$AccountingList = @()
$OtherList = @()

Write-Output "Sorting users into US, UK, and Accounting lists..."
foreach ($user in $InactiveUserList) {
    $isOther = $false
    if($user.Department -like "*Acc*" -or $user.JobTitle -like "*Accountant*" -or $user.JobTitle -like "*Accounting*") {
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
$FilenameAll = "InactiveUsers " + $ReportTimestamp + ".xlsx"
Try {
	$InactiveUserList | Export-Excel -Path $FilenameAll -WorksheetName "ALL Inactive Users"
	$USList | Export-Excel -Path $FilenameAll -WorksheetName "US Inactive"
	$UKList | Export-Excel -Path $FilenameAll -WorksheetName "UK Inactive"
	$AccountingList | Export-Excel -Path $FilenameAll -WorksheetName "Accounting Inactive"
	$OtherList | Export-Excel -Path $FilenameAll -WorksheetName "Unknown Inactive"
} Catch {
	Write-Error -Message $_.Exception.Message
	Write-Output "Error exporting user lists to $FilenameAll."
}

# File to send to US team
$FilenameUSFP = "US InactiveUsers " + $ReportTimestamp + ".xlsx"
Try {
	$USList | Export-Excel -Path $FilenameUSFP -WorksheetName "US Inactive"
	$OtherList | Export-Excel -Path $FilenameUSFP -WorksheetName "Unknown Inactive"
} Catch {
	Write-Error -Message $_.Exception.Message
	Write-Output "Error exporting user lists to $FilenameUSFP."
}

# File to send to UK team
$FilenameUKFP = "UK InactiveUsers " + $ReportTimestamp + ".xlsx"
Try {
	$UKList | Export-Excel -Path $FilenameUKFP -WorksheetName "UK Inactive"
	$OtherList | Export-Excel -Path $FilenameUKFP -WorksheetName "Unknown Inactive"
} Catch {
	Write-Error -Message $_.Exception.Message
	Write-Output "Error exporting user lists to $FilenameUKFP."
}

# File to send to Accounting team
$FilenameAcct = "Acct InactiveUsers " + $ReportTimestamp + ".xlsx"
Try {
	$AccountingList | Export-Excel -Path $FilenameAcct -WorksheetName "Accounting Inactive"
	$OtherList | Export-Excel -Path $FilenameAcct -WorksheetName "Unknown Inactive"
} Catch {
	Write-Error -Message $_.Exception.Message
	Write-Output "Error exporting user lists to $FilenameAcct."
}


# Upload exported Excel document to Azure blob storage
$Filelist = $FilenameAll, $FilenameUSFP, $FilenameUKFP, $FilenameAcct
foreach ($file in $Filelist) {
	Write-Output "Uploading $file to wbprodoffice365 storage account in o365-reports container..."
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


Disconnect-MgGraph
Disconnect-ExchangeOnline -Confirm:$false