# Description: 
# Checks if any IIS worker processes exceeds server CPU usage threshold of 90%. 
# Runs for three consecutive iterations and sends email reporting CPU usage of 
# top 5 worker processes if threshold exceeded.

function sendEmail {
    Param([string]$subject, [string]$body)

    $EmailFrom = "email@domain.com"
    # change to desired recipient
    $EmailTo = "email@domain.com"
    $EmailBody = $body
    $EmailSubject = $subject
    $SMTPServer = "smtpServer"

    Send-MailMessage -From $EmailFrom -To $EmailTo -Priority High -Subject $EmailSubject -BodyAsHtml $EmailBody -SmtpServer $SMTPServer
    "Emailed $subject to $EmailTo"
}

function printProcessAsTableRow {
    Param ([array]$processArray)

    $string = ""

    foreach($process in $processTable) {

        $AppPoolName = $process.AppPoolName
        $ProcessID   = $process.IDProcess
        $ProcessCPU  = $process.PercentCPU

        $string += "<tr>
                        <td style='border: 1px solid black; text-align: center; padding: 5px'>$AppPoolName</td>
                        <td style='border: 1px solid black; text-align: center; padding: 5px'>$ProcessID</td>
                        <td style='border: 1px solid black; text-align: center; padding: 5px'>$ProcessCPU</td>"
        $string += "</tr>"
    }
    return $string
}

function generateEmailMessage {
    Param([array]$processTable, [hashtable]$appPoolExceed)

    $currentDate = Get-Date -Format "MM/dd/yyyy HH:mm tt"

    $message = "Time of event: <b>$currentDate</b><br>Source: $server<br>Alert description: 1 or more app pool(s) exceeded the CPU Usage threshold of 87% for $totalProcessExceed consecutive iterations running at 1 minute intervals."
    
    foreach ($appPool in $appPoolExceed.GetEnumerator()) {
        $message += "<br>The app pool <b>$($appPool.Name)</b> reported <span style='background-color: yellow'>&nbsp;<b>$($appPool.Value)% CPU</b>&nbsp;</span> usage during an iteration."
    }

    $processExceed = printProcessAsTableRow -processArray $processTable

    $message = $message + "
        <br><br><b>The below table shows the top 5 most recent worker processes that are currently running.</b>
        <table style='border: 1px solid black'>
            <tr>
                <th style='border: 1px solid black; text-align: center; padding: 5px'>App Pool Name</th>
                <th style='border: 1px solid black; text-align: center; padding: 5px'>Process ID</th>
                <th style='border: 1px solid black; text-align: center; padding: 5px'>CPU %</th>
            </tr> "

    $message = $message + $processExceed + "</table>"

    return $message
}

# Replace PID key on hashtable with App pool name for easy readability
function Get-AppPoolName {
    Param([hashtable]$appPoolTable)

    $workerprocess = Get-WMIObject Win32_PerfFormattedData_PerfProc_Process|Where Name -Like "w3wp*" | Sort-Object -Property PercentProcessorTime -Descending |Select IDProcess

    foreach($processID in $workerprocess.IDProcess) {
        foreach($appPool in $appPoolTable.Keys 2>$null) {
            if($appPool -eq $processID) {
                $processName = (Get-WmiObject Win32_Process -Filter "processid = $processID" | Select CommandLine).CommandLine.split('(?=(?[^"]|"[^"]*")*$)')[1]
                $appPoolTable.$processName = $appPoolTable.$appPool
                $appPoolTable.Remove($appPool)
            }
        }
    }
    return $appPoolTable
}

function main {
    # Server name
    $server = [System.Net.Dns]::GetHostName().ToUpper()
    $totalProcessExceed = 0
    $appPoolExceed = @{}
    $cpuThreshold = 87
    $NumberOfLogicalProcessors=(Get-WmiObject -class Win32_processor | Measure-Object -Sum NumberOfLogicalProcessors).Sum

    # Check if any worker process exceeds threshold over 3 iterations of 1 minute each interval
    for($i = 0;$i -lt 3;$i++) {
        Write-Host "Iteration: $i"
        # Get table of worker processes with highest CPU % first
        $CPUMemory = Get-WMIObject Win32_PerfFormattedData_PerfProc_Process|Where Name -Like "w3wp*" | Sort-Object -Property PercentProcessorTime -Descending |
        Select -first 5 Name,IDProcess,PercentProcessorTime, @{Name="CPU %";e={[math]::Round($_.PercentProcessorTime/$NumberOfLogicalProcessors)}}
        Write-Output $CPUMemory

        # Find if any worker process exceeds threshold
        foreach ($process in $CPUMemory) {
            if ($process.'CPU %' -gt $cpuThreshold) {
                $totalProcessExceed++

                # Record instace of app pool and memory usage in another table
                if( -not $appPoolExceed.ContainsKey($process.IDProcess)) {
                    $appPoolExceed.Add($process.IDProcess, $process.'CPU %')
                }

                break
            }
        }
        Write-Host "Sleeping for 60 seconds..."
        Start-Sleep -Seconds 60
    }

    # If exceed threshold, send email to WebAdmin
    if($totalProcessExceed -gt 0) {
        # Gets app pool name from process ID and combines table from $CPUMemory
        $processTable = @()
        foreach ($process in $CPUMemory) {
            $processID = $process.IDProcess
            # Gets app pool name from process details
            $processName = (Get-WmiObject Win32_Process -Filter "processid = $processID" | Select CommandLine).CommandLine.split('(?=(?[^"]|"[^"]*")*$)')[1]

            $processTable += [PSCustomObject]@{
                AppPoolName = $processName
                ProcessName = $process.Name
                IDProcess = $processID
                PercentProcessorTime = $process.PercentProcessorTime
                PercentCPU = $process.'CPU %'
            }

            # Change PID to App pool name for easy readability
            $appPoolExceed = Get-AppPoolName -appPoolTable $appPoolExceed
        }

        # Send email to webadmin
        $message = generateEmailMessage -processTable $processTable -appPoolExceed $appPoolExceed
        Write-Output $message
        sendEmail -subject "$server || CPU Usage Exceeds Threshold of 87%" -body $message
    }
    else {
        Write-Output "No App Pools exceeding CPU usage."
    }
}

main