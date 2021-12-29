<#
.NAME
    Create Azure DevOps Build & Release Pipelines
.SYNOPSIS
    Creates new build and release pipelines for an existing Azure DevOps project based on GUI input.
.NOTES
    Last updated 9/8/2021, AP.
#>

#---------------[Form]-------------------

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$form                            = New-Object system.Windows.Forms.Form
$form.ClientSize                 = New-Object System.Drawing.Point(495,475) #495,430
$form.text                       = "Create Azure DevOps Build-Release Pipelines"
$form.TopMost                    = $false
$form.StartPosition              = 'CenterScreen'

$projectName                     = New-Object system.Windows.Forms.TextBox
$projectName.multiline           = $false
$projectName.width               = 204
$projectName.height              = 20
$projectName.location            = New-Object System.Drawing.Point(116,20)
$projectName.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$projectNameLabel                = New-Object system.Windows.Forms.Label
$projectNameLabel.text           = "Project Name :"
$projectNameLabel.AutoSize       = $true
$projectNameLabel.width          = 25
$projectNameLabel.height         = 10
$projectNameLabel.location       = New-Object System.Drawing.Point(21,24)
$projectNameLabel.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$appTypeLabel                    = New-Object system.Windows.Forms.Label
$appTypeLabel.text               = "App Type :"
$appTypeLabel.AutoSize           = $true
$appTypeLabel.width              = 25
$appTypeLabel.height             = 10
$appTypeLabel.location           = New-Object System.Drawing.Point(21,60)
$appTypeLabel.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$appTypeBox                         = New-Object system.Windows.Forms.ComboBox
$appTypeBox.text                    = "Select app type"
$appTypeBox.width                   = 204
$appTypeBox.height                  = 20
@('NonCompiled','Merged .NETCore-React','React') | ForEach-Object {[void] $appTypeBox.Items.Add($_)}
$appTypeBox.location                = New-Object System.Drawing.Point(116,56)
$appTypeBox.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$deployToLabel                   = New-Object system.Windows.Forms.Label
$deployToLabel.text              = "Deploy To :"
$deployToLabel.AutoSize          = $true
$deployToLabel.width             = 25
$deployToLabel.height            = 10
$deployToLabel.location          = New-Object System.Drawing.Point(21,98)
$deployToLabel.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$deployType                      = New-Object system.Windows.Forms.ComboBox
$deployType.text                 = "Select deployment"
$deployType.width                = 204
$deployType.height               = 20
@('Azure','IntranetServer1','IntranetServer2','InternetServer') | ForEach-Object {[void] $deployType.Items.Add($_)}
$deployType.location             = New-Object System.Drawing.Point(116,94)
$deployType.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$createButton                    = New-Object system.Windows.Forms.Button
$createButton.text               = "Create"
$createButton.width              = 94
$createButton.height             = 80
$createButton.location           = New-Object System.Drawing.Point(360,28)
$createButton.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$appPathLabel                = New-Object system.Windows.Forms.Label
$appPathLabel.text           = "App File Path :"
$appPathLabel.AutoSize       = $true
$appPathLabel.width          = 25
$appPathLabel.height         = 10
$appPathLabel.location       = New-Object System.Drawing.Point(21,136)
$appPathLabel.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$appPath                     = New-Object system.Windows.Forms.TextBox
$appPath.multiline           = $false
$appPath.width               = 339
$appPath.height              = 20
$appPath.location            = New-Object System.Drawing.Point(116,132)
$appPath.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$pat                     = New-Object system.Windows.Forms.TextBox
$pat.multiline           = $false
$pat.width               = 339
$pat.height              = 20
$pat.location            = New-Object System.Drawing.Point(116,170)
$pat.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$patLabel                = New-Object system.Windows.Forms.Label
$patLabel.text           = "VSTS PAT :"
$patLabel.AutoSize       = $true
$patLabel.width          = 25
$patLabel.height         = 10
$patLabel.location       = New-Object System.Drawing.Point(21,174)
$patLabel.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


$logLabel                        = New-Object system.Windows.Forms.Label
$logLabel.text                   = "Log :"
$logLabel.AutoSize               = $true
$logLabel.width                  = 25
$logLabel.height                 = 10
$logLabel.location               = New-Object System.Drawing.Point(21,230)
$logLabel.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$outputBox                       = New-Object system.Windows.Forms.TextBox
$outputBox.multiline             = $true
$outputBox.ReadOnly              = $true
$outputBox.WordWrap              = $false
$outputBox.ScrollBars            = "Both"
$outputBox.width                 = 442
$outputBox.height                = 197
$outputBox.location              = New-Object System.Drawing.Point(21,250)
$outputBox.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form.controls.AddRange(@($projectName,$projectNameLabel,$appTypeLabel,$appTypeBox,$deployToLabel,$deployType,$createButton,$appPathLabel,$appPath,$pat,$patLabel,$outputBox,$logLabel))


#------------[Functions]------------

$BldJsonTemplateDirectory = ".\JSON Pipeline Templates\Build\"
$RlsJsonTemplateDirectory = ".\JSON Pipeline Templates\Release\"
$organizationName = "orgName"

# Requires VSTS PAT with build and release read, execute, & manage access
function Get-Header {
    $token = $pat.Text
    $authentication = [Text.Encoding]::ASCII.GetBytes(":$token")
    $authentication = [System.Convert]::ToBase64String($authentication)
    $headers = @{
        Authorization = ("Basic {0}" -f $authentication)
    }
    return $headers
}

# Gets project ID for creating build definition
function Get-ProjectID($projectName) {
    $Url = "https://dev.azure.com/" + $organizationName + "/_apis/projects/" + $projectName + "?api-version=6.0"
    $headers = Get-Header
    $projectInfo = Invoke-RestMethod -Method Get -Uri $Url -Headers $headers
    return $projectInfo.id
}

# Gets project's existing service connection(s) name, ID, and type
function Get-ServiceConnection($projectName) {
    $Url = "https://dev.azure.com/" + $organizationName + "/" + $projectName + "/_apis/serviceendpoint/endpoints?api-version=6.0-preview.4"
    $headers = Get-Header
    $connList = Invoke-RestMethod -Method Get -Uri $Url -Headers $headers
    $connInfo = $connList.value | Select-Object name, id, type
    return $connInfo
}

# Gets values for project's build pipeline ID(s)
function Get-BuildID($projectName) {
    $Url = "https://" + $organizationName + "/" + $projectName + "/_apis/build/definitions?api-version=5.1"
    $headers = Get-Header
    $buildList = Invoke-RestMethod -Method Get -Uri $Url -Headers $headers
    return $buildList.value.id
}

function New-ReleaseDef($projectId, $dev) {
    $outputBox.Text = $outputBox.Text + "`r`nCreating release definition(s)..."

    $azureConn = Get-ServiceConnection $projectName.Text | Where-Object {$_.type -eq "azurerm"}
    $buildId = Get-BuildID $projectName.Text
    $headers = Get-Header
    $Url = "https://vsrm.dev.azure.com/" + $organizationName + "/" + $projectName.Text + "/_apis/release/definitions?api-version=6.0"

    for(($i = 0); $i -le $dev; $i++) {
        switch($deployType.SelectedItem) {
            'IntranetServer1' {
                $releaseDefTemplate = $RlsJsonTemplateDirectory + "PROD-intranet1Deploy.json"

                # Modify template to project's information
                $releaseDef = Get-Content $releaseDefTemplate | ConvertFrom-Json 
                $releaseDef.name = $projectName.Text + ".prod"
                $releaseDef.environments[0].deployPhases[0].workflowTasks[0].inputs.WebsiteName = "intranetServer1/" + $appPath.Text
                if($appTypeBox.SelectedItem -ne 'React') {
                    $releaseDef.environments[0].deployPhases[0].workflowTasks[0].inputs.WebDeployPackage += $projectName.Text + ".zip"
                    $releaseDef.environments[0].deployPhases[0].workflowTasks[0].name += $projectName.Text + ".zip" 
                }
            }
            'IntranetServer2' {
                $releaseDefTemplate = $RlsJsonTemplateDirectory + "PROD-intranet2Deploy.json"
                $releaseDefDevTemplate = $RlsJsonTemplateDirectory + "DEV-intranet2Deploy.json"

                # Modify template to project's information
                if ($i -eq 1) { 
                    $releaseDef = Get-Content $releaseDefDevTemplate | ConvertFrom-Json
                    $releaseDef.environments[0].deployPhases[0].workflowTasks[0].inputs.WebsiteName = "dev-intranetServer2/" + $appPath.Text
                } 
                else {
                    $releaseDef = Get-Content $releaseDefTemplate | ConvertFrom-Json 
                    $releaseDef.name = $projectName.Text + ".prod"
                    $releaseDef.environments[0].deployPhases[0].workflowTasks[0].inputs.WebsiteName = "intranetServer2/" + $appPath.Text
                }
                if($appTypeBox.SelectedItem -ne 'React') {
                    $releaseDef.environments[0].deployPhases[0].workflowTasks[0].inputs.WebDeployPackage += $projectName.Text + ".zip"
                    $releaseDef.environments[0].deployPhases[0].workflowTasks[0].name += $projectName.Text + ".zip" 
                }
            }
            'InternetServer' {
                $releaseDefTemplate = $RlsJsonTemplateDirectory + "PROD-internetDeploy.json"
                $releaseDefDevTemplate = $RlsJsonTemplateDirectory + "DEV-internetDeploy.json"

                # Modify template to project's information
                if ($i -eq 1) { 
                    $releaseDef = Get-Content $releaseDefDevTemplate | ConvertFrom-Json 
                    $releaseDef.environments[0].deployPhases[0].workflowTasks[0].inputs.WebsiteName = "dev-internetServer/" + $appPath.Text
                    if($appTypeBox.SelectedItem -ne 'React') {
                        $releaseDef.environments[0].deployPhases[0].workflowTasks[0].inputs.WebDeployPackage += $projectName.Text + ".zip"
                        $releaseDef.environments[0].deployPhases[0].workflowTasks[0].name += $projectName.Text + ".zip"
                    }
                } 
                else { 
                    $releaseDef = Get-Content $releaseDefTemplate | ConvertFrom-Json 
                    $releaseDef.name = $projectName.Text + ".prod"
                    $releaseDef.environments[0].deployPhases[0].workflowTasks[1].inputs.WebsiteName = "internetServer/" + $appPath.Text
                    if($appTypeBox.SelectedItem -ne 'React') {
                        $releaseDef.environments[0].deployPhases[0].workflowTasks[0].inputs.SourcePath += $projectName.Text + ".zip"
                        $releaseDef.environments[0].deployPhases[0].workflowTasks[0].name += $projectName.Text + ".zip"
                        $releaseDef.environments[0].deployPhases[0].workflowTasks[1].inputs.WebDeployPackage += $projectName.Text + ".zip"
                        $releaseDef.environments[0].deployPhases[0].workflowTasks[1].name += $projectName.Text + ".zip"
                    }
                }
            }
            'Azure' {
                $releaseDefTemplate = $RlsJsonTemplateDirectory + "PROD-AzureDeploy.json"
                $releaseDefDevTemplate = $RlsJsonTemplateDirectory + "DEV-AzureDeploy.json"

                # Modify template to project's information
                if ($i -eq 1) { 
                    $releaseDef = Get-Content $releaseDefDevTemplate | ConvertFrom-Json 
                    $releaseDef.environments[0].deployPhases[0].workflowTasks[1].inputs.Package = "$" + "(System.DefaultWorkingDirectory)/_" + $projectName.Text + ".dev/drop/*.zip"
                    $releaseDef.environments[0].deployPhases[0].workflowTasks[1].inputs.ConnectedServiceName = $azureConn.id
                    $releaseDef.environments[0].deployPhases[0].workflowTasks[0].inputs.appName = $projectName.Text + "-dev"
                    $releaseDef.environments[0].deployPhases[0].workflowTasks[0].name += $projectName.Text + "-dev"
                    $releaseDef.environments[0].deployPhases[0].workflowTasks[1].inputs.WebAppName = $projectName.Text + "-dev"
                    $releaseDef.environments[0].deployPhases[0].workflowTasks[1].name += $projectName.Text + "-dev"
                } 
                else { 
                    $releaseDef = Get-Content $releaseDefTemplate | ConvertFrom-Json 
                    $releaseDef.name = $projectName.Text + ".prod"
                    $releaseDef.environments[0].deployPhases[0].workflowTasks[0].inputs.Package = "$" + "(System.DefaultWorkingDirectory)/_" + $projectName.Text + ".master/drop/*.zip"
                    $releaseDef.environments[0].deployPhases[0].workflowTasks[0].inputs.WebAppName = $projectName.Text
                    $releaseDef.environments[0].deployPhases[0].workflowTasks[0].name += $projectName.Text
                }

                $releaseDef.environments[0].deployPhases[0].workflowTasks[0].inputs.ConnectedServiceName = $azureConn.id
            }
            Default { 
                $outputBox.Text = $outputBox.Text + "`r`nDeploy Type invalid." 
            }
        }
        
        # If $dev = 1, then also create dev pipeline
        if($i -eq 1) {
            $releaseDef.name = $projectName.Text + ".dev"
            $releaseDef.artifacts[0].alias = "_" + $projectName.Text + ".dev"
            $releaseDef.artifacts[0].definitionReference.definition.name = $projectName.Text + ".dev"
            $releaseDef.artifacts[0].definitionReference.definition.id = $buildId[1]
            $releaseDef.artifacts[0].sourceId = $projectId + ":" + $buildId[1]
        }
        else {
            $releaseDef.artifacts[0].alias = "_" + $projectName.Text + ".master"
            $releaseDef.artifacts[0].definitionReference.definition.name = $projectName.Text + ".master"
            $releaseDef.artifacts[0].definitionReference.definition.id = $buildId[0]
            $releaseDef.artifacts[0].sourceId = $projectId + ":" + $buildId[0]
        }

        $releaseDef.artifacts[0].definitionReference.project.id = $projectId
        $releaseDef.artifacts[0].definitionReference.project.name = $projectName.Text
        $releaseDef.triggers[0].artifactAlias = $releaseDef.artifacts[0].alias

        # Convert back to Json
        $releaseDefJson = @($releaseDef) | ConvertTo-Json -Depth 99
    
        try { 
            Invoke-RestMethod -Uri $Url -Method Post -Body $releaseDefJson -ContentType "application/json" -Headers $headers
            $outputBox.Text = $outputBox.Text + "`r`nSuccessfully created release definition " + $releaseDef.name + " for project " + $projectName.Text
        } catch {
            $outputBox.Text = $outputBox.Text + "`r`nAn error occurred trying to create release definition " + $releaseDef.name
            $outputBox.Text = $outputBox.Text + "`r`nPROJECT: " + $projectName.Text + "`r`n" + $_ + "`r`n"
        }
    }
}

function New-BuildDef {
    $outputBox.Text = "Creating build definition(s)..."

    $projectID = Get-ProjectID $projectName.Text
    $githubConn = Get-ServiceConnection $projectName.Text | Where-Object {$_.type -eq "github"}
    $Url = "https://" + $organizationName + "/" + $projectName.Text + "/_apis/build/definitions?api-version=6.0"
    $headers = Get-Header
    $dev = 0 # Value is default 0 to skip dev pipeline

    # Determine app type based on user selection
    if($appTypeBox.SelectedItem -eq 'Merged .NETCore-React' -and $deployType.SelectedItem -eq 'Azure') {
        $appType = "Core-React-Azure"
    }
    elseif($appTypeBox.SelectedItem -eq 'Merged .NETCore-React' -and ($deployType.SelectedItem -match 'Intranet' -or $deployType.SelectedItem -eq 'InternetServer')) {
        $appType = "Core-React-Internal"
    }
    else { $appType = $appTypeBox.SelectedItem }

    # Gets build definition JSON template for corresponding app type
    switch($appType) {
        'NonCompiled' {
            $buildDefTemplate = $BldJsonTemplateDirectory + "MASTER-NonCompiled.json"
            $buildDef = Get-Content $buildDefTemplate | ConvertFrom-Json

            # Modify template to project's information
            $buildDef.process.phases.steps[1].inputs.archiveFile = "$" + "(Build.ArtifactStagingDirectory)/" + $projectName.Text +".zip"
        }
        'Core-React-Azure' {
            $dev = 1 # dev set to 1 to also create dev pipeline later
            $buildDefTemplate = $BldJsonTemplateDirectory + "MASTER-Core-React-AzurePublish.json"
            $buildDef = Get-Content $buildDefTemplate | ConvertFrom-Json
            $buildDefDevTemplate = $BldJsonTemplateDirectory + "DEV-Core-React-AzurePublish.json"
        }
        'Core-React-Internal' {
            $dev = 1 # dev set to 1 to also create dev pipeline later
            $buildDefTemplate = $BldJsonTemplateDirectory + "MASTER-Core-React-internalPublish.json"
            $buildDef = Get-Content $buildDefTemplate | ConvertFrom-Json
            $buildDefDevTemplate = $BldJsonTemplateDirectory + "DEV-Core-React-internalPublish.json"
        }
        'React' {
            $dev = 1 # dev set to 1 to also create dev pipeline later
            $buildDefTemplate = $BldJsonTemplateDirectory + "MASTER-React.json"
            $buildDef = Get-Content $buildDefTemplate | ConvertFrom-Json
            $buildDefDevTemplate = $BldJsonTemplateDirectory + "DEV-React.json"
        }
        Default { 
            $appType = 0
            $outputBox.Text = $outputBox.Text + "`r`nApp Type invalid." 
        }
    }
 
    if($appType -eq 0) {
        $outputBox.Text = $outputBox.Text + "`r`nCannot create build pipeline. Please check your input and try again."
    }
    else { # Make project specific changes to JSON template and send POST request to create build definition(s)
        for(($i = 0); $i -le $dev; $i++) {
            $buildDef.name = $projectName.Text + ".master"
            
            # If $dev = 1, then also create dev pipeline
            if($i -eq 1) {
                $buildDef = Get-Content $buildDefDevTemplate | ConvertFrom-Json
                $buildDef.name = $projectName.Text + ".dev"
            }

            $buildDef.project.id = $projectID
            $buildDef.repository.id = $buildDef.repository.name = "githubOrgName/" + $projectName.Text
            $buildDef.repository.url = "https://github.com/githubOrgName/" + $projectName.Text + ".git"
            $buildDef.repository.properties.apiUrl = "https://api.github.com/repos/githubOrgName/" + $projectName.Text
            $buildDef.repository.properties.connectedServiceId = $githubConn.id

            # Convert back to Json for request body
            $buildDefJson = @($buildDef) | ConvertTo-Json -Depth 99

            try { 
                Invoke-RestMethod -Uri $Url -Method Post -Body $buildDefJson -ContentType "application/json" -Headers $Headers
                $outputBox.Text = $outputBox.Text + "`r`nSuccessfully created build definition " + $buildDef.name + " for project " + $projectName.Text
                $buildSuccess = 1
            } catch {
                $outputBox.Text = $outputBox.Text + "`r`nAn error occurred trying to create build definition " + $buildDef.name
                $outputBox.Text = $outputBox.Text + "`r`nPROJECT: " + $projectName.Text + "`r`n" + $_ + "`r`n"
                $buildSuccess = 0
            }
        }
    }

    # Create release pipeline if build was successful
    if($buildSuccess -eq 1) {
        New-ReleaseDef $projectId $dev
        $outputBox.Text = $outputBox.Text + "`r`nNOTE: BEFORE RUNNING PIPELINES"
        $outputBox.Text = $outputBox.Text + "`r`nIf branch names on GH repo is not using `"dev`" or `"master`", you need to change the branch specification in build pipeline under CI triggers."
        $outputBox.Text = $outputBox.Text + "`r`nSet agent pool on release pipeline tasks."
        $outputBox.Text = $outputBox.Text + "`r`nSet password variable on release pipeline."
    }
    else {
        $outputBox.Text = $outputBox.Text + "`r`nCannot create release pipeline because build failed."
    }
    $outputBox.Text = $outputBox.Text + "`r`nFinished.`r`n"
}

#-------------[Script]--------------

$createButton.Add_Click({ New-BuildDef })

#------------[Show form]------------

[void]$Form.ShowDialog()