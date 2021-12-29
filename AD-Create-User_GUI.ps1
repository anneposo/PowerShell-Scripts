<#
.NAME
    AD Account Creation Tool
.SYNOPSIS
    Creates new AD user based on GUI input.
.DESCRIPTION
    Creates new AD user based on GUI input. User is created with default group permissions (DomainUsers, InternetUsers) in specified OU.
    After new AD user account is created, an NT & Email Account Log excel sheet is also updated with the user/request information.
    First checks if an account already exists on AD with the input username. If user does not exist, then proceeds to create new AD account.
    Telephone and employee number fields can be left blank if information is unavailable at time of creation, all other fields are required.
    For consultants: 
        OU is set to OU=Consultant, 
        user is not added to InternetUsers group.
        display name is appended with (Consultant), 
        account expiration date is set to 1 year from current date,
        places "Consultant" in Telephone notes section of AD user.
.NOTES
    Last updated 9/6/2021, AP.
    After script creates account, sometimes it can take a few minutes until user appears on search in Active Directory,
    but the user is still searchable via PowerShell with command "Get-ADUser <username>"
#>

# Import active directory module for running AD cmdlets
Import-Module activedirectory

#---------------[Form]-------------------

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$form                            = New-Object system.Windows.Forms.Form
$form.ClientSize                 = New-Object System.Drawing.Point(490,470)
$form.text                       = "Create New AD Account"
$form.TopMost                    = $false
$form.StartPosition              = 'CenterScreen'

$firstNameBox                     = New-Object system.Windows.Forms.TextBox
$firstNameBox.multiline           = $false
$firstNameBox.width               = 125
$firstNameBox.height              = 20
$firstNameBox.location            = New-Object System.Drawing.Point(100,20)
$firstNameBox.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$firstNameLabel                = New-Object system.Windows.Forms.Label
$firstNameLabel.text           = "First Name:"
$firstNameLabel.AutoSize       = $true
$firstNameLabel.width          = 25
$firstNameLabel.height         = 10
$firstNameLabel.location       = New-Object System.Drawing.Point(21,23)
$firstNameLabel.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lastNameBox                     = New-Object system.Windows.Forms.TextBox
$lastNameBox.multiline           = $false
$lastNameBox.width               = 125
$lastNameBox.height              = 20
$lastNameBox.location            = New-Object System.Drawing.Point(322,20)
$lastNameBox.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lastNameLabel                = New-Object system.Windows.Forms.Label
$lastNameLabel.text           = "Last Name:"
$lastNameLabel.AutoSize       = $true
$lastNameLabel.width          = 25
$lastNameLabel.height         = 10
$lastNameLabel.location       = New-Object System.Drawing.Point(241,23)
$lastNameLabel.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$usernameBox                     = New-Object system.Windows.Forms.TextBox
$usernameBox.multiline           = $false
$usernameBox.width               = 125
$usernameBox.height              = 20
$usernameBox.location            = New-Object System.Drawing.Point(100,56)
$usernameBox.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$usernameLabel                = New-Object system.Windows.Forms.Label
$usernameLabel.text           = "Username:"
$usernameLabel.AutoSize       = $true
$usernameLabel.width          = 25
$usernameLabel.height         = 10
$usernameLabel.location       = New-Object System.Drawing.Point(21,59)
$usernameLabel.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$passwordBox                     = New-Object system.Windows.Forms.TextBox
$passwordBox.multiline           = $false
$passwordBox.width               = 125
$passwordBox.height              = 20
$passwordBox.location            = New-Object System.Drawing.Point(322,56)
$passwordBox.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$passwordLabel                = New-Object system.Windows.Forms.Label
$passwordLabel.text           = "Password:"
$passwordLabel.AutoSize       = $true
$passwordLabel.width          = 25
$passwordLabel.height         = 10
$passwordLabel.location       = New-Object System.Drawing.Point(241,59)
$passwordLabel.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$divisionLabel                   = New-Object system.Windows.Forms.Label
$divisionLabel.text              = "Division:"
$divisionLabel.AutoSize          = $true
$divisionLabel.width             = 25
$divisionLabel.height            = 10
$divisionLabel.location          = New-Object System.Drawing.Point(21,135)
$divisionLabel.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$division                      = New-Object system.Windows.Forms.ComboBox
$division.text                 = "Select option.."
$division.width                = 125
$division.height               = 20
@('ADM', 'AVI', 'BFM', 'BRCD', 'BSD', 'CGRG', 'CSG', 'DES', 'DSG', 'EPD', 'FIS', 'FLT', 'GMED', 'HRD', `
'IAG', 'ITDOI', 'ITDSA', 'LDD', 'OSD', 'PMDI', 'PMDII', 'PMDIII', 'RMD', 'RMO', 'SMD', 'SMP', `
'SPSO', 'SWED', 'SWMD', 'SWPD', 'SWQD', 'TPP', 'TSM', 'WSD', 'WWD') | ForEach-Object {[void] $division.Items.Add($_)}
$division.location             = New-Object System.Drawing.Point(100,132)
$division.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$consultantLabel                   = New-Object system.Windows.Forms.Label
$consultantLabel.text              = "Consultant:"
$consultantLabel.AutoSize          = $true
$consultantLabel.width             = 25
$consultantLabel.height            = 10
$consultantLabel.location          = New-Object System.Drawing.Point(241,135)
$consultantLabel.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$isConsultant                      = New-Object system.Windows.Forms.ComboBox
$isConsultant.text                 = "Select option.."
$isConsultant.width                = 125
$isConsultant.height               = 20
@('Yes', 'No') | ForEach-Object {[void] $isConsultant.Items.Add($_)}
$isConsultant.location             = New-Object System.Drawing.Point(322,132)
$isConsultant.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$createButton                    = New-Object system.Windows.Forms.Button
$createButton.text               = "Create User"
$createButton.width              = 428
$createButton.height             = 35
$createButton.location           = New-Object System.Drawing.Point(21,210)
$createButton.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$telephoneLabel                = New-Object system.Windows.Forms.Label
$telephoneLabel.text           = "Telephone:"
$telephoneLabel.AutoSize       = $true
$telephoneLabel.width          = 25
$telephoneLabel.height         = 10
$telephoneLabel.location       = New-Object System.Drawing.Point(21,97)
$telephoneLabel.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$telephoneBox                     = New-Object system.Windows.Forms.TextBox
$telephoneBox.multiline           = $false
$telephoneBox.width               = 125
$telephoneBox.height              = 20
$telephoneBox.location            = New-Object System.Drawing.Point(100,94)
$telephoneBox.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$employeeNum                     = New-Object system.Windows.Forms.TextBox
$employeeNum.multiline           = $false
$employeeNum.width               = 125
$employeeNum.height              = 20
$employeeNum.location            = New-Object System.Drawing.Point(322,94)
$employeeNum.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$employeeNumLabel                = New-Object system.Windows.Forms.Label
$employeeNumLabel.text           = "Employee #:"
$employeeNumLabel.AutoSize       = $true
$employeeNumLabel.width          = 25
$employeeNumLabel.height         = 10
$employeeNumLabel.location       = New-Object System.Drawing.Point(241,97)
$employeeNumLabel.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$requestorLabel                = New-Object system.Windows.Forms.Label
$requestorLabel.text           = "Requestor:"
$requestorLabel.AutoSize       = $true
$requestorLabel.width          = 25
$requestorLabel.height         = 10
$requestorLabel.location       = New-Object System.Drawing.Point(21,173)
$requestorLabel.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$requestorBox                     = New-Object system.Windows.Forms.TextBox
$requestorBox.multiline           = $false
$requestorBox.width               = 125
$requestorBox.height              = 20
$requestorBox.location            = New-Object System.Drawing.Point(100,170)
$requestorBox.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$requestNum                     = New-Object system.Windows.Forms.TextBox
$requestNum.multiline           = $false
$requestNum.width               = 125
$requestNum.height              = 20
$requestNum.location            = New-Object System.Drawing.Point(322,170)
$requestNum.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$requestNumLabel                = New-Object system.Windows.Forms.Label
$requestNumLabel.text           = "Request #:"
$requestNumLabel.AutoSize       = $true
$requestNumLabel.width          = 25
$requestNumLabel.height         = 10
$requestNumLabel.location       = New-Object System.Drawing.Point(241,173)
$requestNumLabel.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$logLabel                        = New-Object system.Windows.Forms.Label
$logLabel.text                   = "Log:"
$logLabel.AutoSize               = $true
$logLabel.width                  = 25
$logLabel.height                 = 10
$logLabel.location               = New-Object System.Drawing.Point(21,270)
$logLabel.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$outputBox                       = New-Object system.Windows.Forms.TextBox
$outputBox.multiline             = $true
$outputBox.ReadOnly              = $true
$outputBox.WordWrap              = $false
$outputBox.ScrollBars            = "Both"
$outputBox.width                 = 442
$outputBox.height                = 150
$outputBox.location              = New-Object System.Drawing.Point(21,290)
$outputBox.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form.controls.AddRange(@($firstNameBox,$firstNameLabel,$lastNameBox,$lastNameLabel,$divisionLabel,$division,$usernameBox,$usernameLabel, `
$passwordBox,$passwordLabel,$consultantLabel,$isConsultant,$telephoneLabel,$telephoneBox,$employeeNum,$employeeNumLabel,$requestorLabel, `
$requestorBox,$requestNumLabel,$requestNum,$outputBox,$logLabel,$createButton))


#------------[Functions]------------

# Adds new user entry in NT & Email Accounts Log
function Update-ExcelSheet {
    $outputBox.Text = $outputBox.Text + "`r`nUpdating NT & Email Accounts Log.."

    $excelFile = "\\serverName\filepath\NT & Email Accounts Log.xlsx"
    
    # Create excel object
    $excel                = New-Object -ComObject Excel.Application
    $excel.Visible        = $false # Opens excel workbook using MS Office Excel application

    # Opens excel file to specific worksheet
    $workbook  = $excel.Workbooks.Open($excelFile)
    $worksheet = $excel.Worksheets.Item('New NT & Email')
    $outputBox.Text = $outputBox.Text + "`r`nOpening Excel file at $excelFile"

    # Gets last row and adds 1 for next empty row
    $worksheet.activate()
    $nextEmptyRow = ($worksheet.UsedRange.rows.count) + 1

    $Firstname 	= $firstNameBox.Text
    $Lastname 	= $lastNameBox.Text

    # Updates next empty row with new AD user/requestor information.
    for ($i=1; $i -le 8; $i++) {
        switch($i) {
            1  { $worksheet.Cells.Item($nextEmptyRow,$i) = (Get-Date).ToString("M/dd/yyyy") }
            2  { $worksheet.Cells.Item($nextEmptyRow,$i) = $division.SelectedItem }
            3  { 
                if($isConsultant.SelectedItem -eq 'Yes') {
                    $worksheet.Cells.Item($nextEmptyRow,$i) = "$Firstname $Lastname (Consultant)"
                    $worksheet.Cells($nextEmptyRow,$i).Font.Size = 12
                }
                else {
                    $worksheet.Cells.Item($nextEmptyRow,$i) = "$Firstname $Lastname (e"  + $employeeNum.Text  + ")"
                    $worksheet.Cells($nextEmptyRow,$i).Font.Size = 12
                } 
            }
            4  { $worksheet.Cells.Item($nextEmptyRow,$i) = (Get-Date).ToString("M/dd/yyyy") }
            5  { $worksheet.Cells.Item($nextEmptyRow,$i) = $requestorBox.Text }
            6  { $worksheet.Cells.Item($nextEmptyRow,$i) = "Anne Poso" }
            7  { $worksheet.Cells.Item($nextEmptyRow,$i) = "SR #" + $requestNum.Text }
            8  { $worksheet.Cells.Item($nextEmptyRow,$i) = "New Account" }
        }
    }

    # Save and close excel file
    $workbook.Save()
    $workbook.Close()
    $excel.Quit()

    $outputBox.Text = $outputBox.Text + "`r`nUpdated Log and closed Excel file."
}

# Creates new AD user with GUI input
function New-PWUser {
    $outputBox.Text = "Creating new user.."

    # Get account properties from user input
    $Username 	= $usernameBox.Text
    $Password 	= $passwordBox.Text
    $Firstname 	= $firstNameBox.Text
    $Lastname 	= $lastNameBox.Text
    $email      = $usernameBox.Text + "@domain"
    $telephone  = $telephoneBox.Text
    $desc       = $division.SelectedItem
    $ipPhone    = $employeeNum.Text

    # If user is a consultant, set different OU, display name, and account expiration date
    if($isConsultant.SelectedItem -eq 'Yes') {
        $OU = "OU=Consultant,OU=USERS,DC=domaincontroller"
        $DisplayName = "$Lastname, $Firstname (Consultant)"
        $expireDate = (Get-Date).AddYears(1).ToString("MM/dd/yyyy")
    }
    else {
        $OU = "OU=USERS,DC=domaincontroller"
        $DisplayName = "$Lastname, $Firstname"
    }

    # Set user's company based on selected division acronym
    switch($desc) {
        'ADM'   { $company =  "Division Name"  }
        'AVI'   { $company =  "Division Name" }
        'BFM'   { $company =  "Division Name" }
        'BRCD'  { $company =  "Division Name" }
        'BSD'   { $company =  "Division Name" }
        'CGRG'  { $company =  "Division Name" }
        'CSG'   { $company =  "Division Name" }
        'DES'   { $company =  "Division Name" }
        'DSG'   { $company =  "Division Name" }
        'EPD'   { $company =  "Division Name" }
        'FIS'   { $company =  "Division Name" }
        'FLT'   { $company =  "Division Name" }
        'GMED'  { $company =  "Division Name" }
        'HRD'   { $company =  "Division Name" }
        'IAG'   { $company =  "Division Name" }
        'ITDOI' { $company =  "Division Name" }
        'ITDSA' { $company =  "Division Name" }
        'LDD'   { $company =  "Division Name" }
        Default { 
            $company = ""
            $outputBox.Text = $outputBox.Text + "`r`nDivision information invalid." 
        }
    }

    #Check to see if the user already exists in AD
	if (Get-ADUser -F {SamAccountName -eq $Username})
	{
		 #If user does exist, give a warning
		 $outputBox.Text = $outputBox.Text + "`r`nError: A user account with username $Username already exist in Active Directory."
	}
	else
	{
		#User does not exist then proceed to create the new user account
		
        #Account will be created in the OU provided by the $OU variable
        try {
            New-ADUser `
                -SamAccountName $Username `
                -UserPrincipalName "$Username@domain" `
                -Name "$Lastname, $Firstname" `
                -GivenName $Firstname `
                -Surname $Lastname `
                -Enabled $True `
                -DisplayName $DisplayName `
                -Path $OU `
                -Company $company `
                -Description $desc `
                -OfficePhone $telephone `
                -EmailAddress $email `
                -Office $company `
                -Department "Department" `
                -AccountPassword (convertto-securestring $Password -AsPlainText -Force) -ChangePasswordAtLogon $True
            
            $outputBox.Text = $outputBox.Text + "`r`nCreated AD user: $Username"
            $userAcc = 1
        }
        catch {
            $outputBox.Text = $outputBox.Text + "`r`nAn error occurred trying to create the user $Username`r`n" +"$_`r`n"
            $userAcc = 0
        }

        # If AD user was created successfully, set additional account properties.
        if($userAcc -eq 1) {
            if($ipPhone) {
                Set-ADUser -identity $Username -add @{ipphone="$ipphone"} # employee/contractor number
            }

            # If user is a consultant, set account expiration date to 1 year.
            if($isConsultant.SelectedItem -eq 'Yes') {
                Set-ADUser -identity $Username -AccountExpirationDate $expireDate
                Set-ADUser -identity $Username -Replace @{info="Consultant"} # Adds "Consultant" to Telephone notes section
            }
            else { 
                Add-ADGroupMember -Identity InternetUsers -Members $Username
            }

            # Update NT & Email Accounts Log after user account is created.
            Update-ExcelSheet
        }

        $outputBox.Text = $outputBox.Text + "`r`nFinished."
	}
}

#-------------[Script]--------------

$createButton.Add_Click({ New-PWUser })

#------------[Show form]------------

[void]$Form.ShowDialog()