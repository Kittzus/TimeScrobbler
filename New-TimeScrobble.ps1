Function Get-OutlookInBox {
 Add-type -assembly 'Microsoft.Office.Interop.Outlook' | out-null
 $olFolders = 'Microsoft.Office.Interop.Outlook.olDefaultFolders' -as [type]
 $outlook = new-object -comobject outlook.application
 $namespace = $outlook.GetNameSpace('MAPI')
 $folder = $namespace.getDefaultFolder($olFolders::olFolderInBox)
 $folder.items |
 Select-Object -Property Subject, ReceivedTime, Importance, SenderName
}

Function Get-OutlookSent {
 Add-type -assembly 'Microsoft.Office.Interop.Outlook' | out-null
 $olFolders = 'Microsoft.Office.Interop.Outlook.olDefaultFolders' -as [type]
 $outlook = new-object -comobject outlook.application
 $namespace = $outlook.GetNameSpace('MAPI')
 $folder = $namespace.getDefaultFolder($olFolders::olFolderSentMail)
 $folder.items |
 Select-Object -Property Subject, SentOn, Importance, To
}

Function Get-OutlookCalendar { 
 Add-type -assembly 'Microsoft.Office.Interop.Outlook' | out-null 
 $olFolders = 'Microsoft.Office.Interop.Outlook.OlDefaultFolders' -as [type]  
 $outlook = new-object -comobject outlook.application 
 $namespace = $outlook.GetNameSpace('MAPI') 
 $folder = $namespace.getDefaultFolder($olFolders::olFolderCalendar) 
 $folder.items | 
 Select-Object -Property Subject, Start, Duration, Location 
}

Function New-HTMLTable {
    param(
        $inputObj,
        $headerText
    )

    $outputArr = @()
    $outputArr += ('<h2>',$headerText,'</h2>' -join '')
    $outputArr += $inputObj | ConvertTo-HTML -Fragment
    return $outputArr
}

Function Set-Folder {
# Function to check if a folder exists, and create it if not.
	param($path)
			
	If (!(Test-Path "filesystem::$path" -ErrorAction SilentlyContinue)) {
		New-Item $path -Type Directory -Force | Out-Null
	}
}

Function Import-ConfigFile {
    # Quick function to read in simple plaintext config files
    # Evaluate switch will expand any powershell variables (e.g. $test) in the imported strings before setting the value

    param(
        $Path,
        [switch]$Evaluate
    )

    $importConfig = Get-Content $path | Where-Object {($_ -notlike '#*') -and ($_)}
    $importConfig | ForEach {
        $splitVar = $_.Split('=')
        If ($splitVar[1] -like '*,*') {
            [array]$value = $splitVar[1].Split(',').Trim()
            If ($evaluate) {
                $evals = @()
                $value | ForEach {
                    $evals += $ExecutionContext.InvokeCommand.ExpandString("$_")
                }
                $value = $evals
            }
        }
        Else {
            [string]$value = $splitvar[1].Trim()
            If ($evaluate) {
                $value = $ExecutionContext.InvokeCommand.ExpandString("$value")
            }
        }
        Set-Variable -Scope Script -Name $splitVar[0].Trim() -Value $value
    }
}

$reportHeader = @"
    <link rel="stylesheet" src="https://necolas.github.io/normalize.css/latest/normalize.css">
    <style>
        body {
            color: #222;
            font-family: sans-serif;
            font-size: 14px;
            margin: 2% 0;
        }
        h1 {
            font-size: 2em;
            font-weight: normal;
            padding: 0 2%;
        }
        h3 {
            font-size: 1.25em;
            font-weight: normal;
            padding: 0 2%;
        }
        table {
            border-collapse: collapse;
            width: 100%;
        }
        tr:nth-child(even) {
            background: #EEE;
        }
        th {
            border-bottom: 1px solid #999;
            font-weight: normal;
            text-align: left;
        }
        td,
        th {
            padding: .25em;
        }
        td:first-child,
        th:first-child    {
            padding-left: 2%;
        }
        td:last-child,
        th:last-child {
            padding-left: 2%;
        }
        h2 {
            font-size: 1.5em;
            font-weight: normal;
            margin: 1 0 .5%;
            padding: 0 2%;
        }
    </style>
"@


Write-Output "`r`nTimeScrobbler v1.0 - 2016-08-18 - https://github.com/Kittzus/TimeScrobbler`r`n"
# Work out path, import Slack module (TEMP: Until the Groups functionality is added to master branch)
$scriptPath = Split-Path $MyInvocation.MyCommand.Path -Parent

# Read in the User.conf and turn into live variables
Import-ConfigFile -Path $scriptPath\User.conf

# Import PSSlack if a token has been provided
If (!$slackToken) {
    Write-Output 'No Slack user token found in User.conf file. Enter here, or [Return] to skip'
    $slackToken = Read-Host 'Slack User Token'
}
If ($slackToken) {
    $modCheck = Get-Module PSSlack
    If (!$modCheck) {
        Write-Output 'PSSlack module not found in $PSModulePath folders. Please install this from https://github.com/RamblingCookieMonster/PSSlack to enable Slack scrobbling!'
        Clear-Variable slackToken
    }
    Else{
        Import-Module PSSlack -Force | Out-Null
    }
}

# Create the output folder if it doesn't already exist
Set-Folder $outputFld

While (!$doneFlag) {
    Clear-Variable moreFlag -ErrorAction SilentlyContinue
    Write-Output "`r`nPlease set the date-range you`'d like to TimeScrobble"
    Write-Output "IMPORTANT: Date's must be in the unambiguous sortable date format yyyy-MM-dd e.g. 2016-08-13 for 13th August 2016"
    Write-Output '[q] Exit'

    While ($runFlag -ne 'y') {
        Clear-Variable validTimeSpan,runFlag -ErrorAction SilentlyContinue

        While (!$validTimeSpan) {
            Clear-Variable validStart,validEnd -ErrorAction SilentlyContinue

            While (!$validStart) {
                $startDate = Read-Host 'Start Date'
                Switch ($startDate) {
                    'q' {
                        exit
                    }
                    default {
                        Try {
                            $validStart = Get-Date $startDate -ErrorAction Stop
                        }
                        Catch {
                            Write-Output "Invalid date entered. Format must be yyyy-MM-dd.`r`n"
                        }
                    }
                }
            }

            While (!$validEnd) {
                $endDate = Read-Host 'End Date'
                Switch ($endDate) {
                    'q' {
                        exit
                    }
                    default {
                        Try {
                            $validEnd = Get-Date $endDate -ErrorAction Stop
                        }
                        Catch {
                            Write-Output "Invalid date entered. Format must be yyyy-MM-dd.`r`n"
                        }
                    }
                }
            }

            $validTimeSpan = New-TimeSpan -Start $validStart -End $validEnd
            If ($validTimeSpan -like '-*') {
                Write-Output 'Invalid time period entered. Start Date must be BEFORE the End Date.'
                Clear-Variable validTimeSpan
            }
        }

        # Build our array of dates
        $days = [Math]::Ceiling($validTimeSpan.TotalDays)+1
        $dateArr = @()
        1..$days | ForEach-Object {
            $dateArr += $validStart
            $validStart = $validStart.AddDays(1)
        }

        Write-Output "About to TimeScrobble $($dateArr.Count) days.`r`n"
        Write-Output "Start Date: $($dateArr[0])"
        Write-Output "End Date: $($validStart)"

        While(@('y','n') -notcontains $runFlag) {
            $runFlag = Read-Host "`r`nBegin? [y/n]"
            If (@('y','n') -notcontains $runFlag) {
                Write-Output "Invalid Entry!`r`n"
            }
        }
    }

    $folderArr += [Environment]::GetFolderPath('Desktop')
    $folderArr += [Environment]::GetFolderPath('Desktop')
    $downloadPath = Get-ItemProperty 'Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' | Select-Object -ExpandProperty '{374DE290-123F-4565-9164-39C4925E467B}'

    # Build out the data sources for the reports if they don't yet exist
    If (!$inboxArr) {
        Write-Output 'Getting Outlook Inbox - This may take some time... No, seriously. Make a sandwich.'
        $inboxArr = Get-OutlookInbox
    }
    If (!$sentArr) {
        Write-Output 'Getting Outlook Sent Items - Could be a while...'
        $sentArr = Get-OutlookSent
    }
    If (!$calArr) {
        Write-Output 'Getting Outlook Calendar - Hopefully wont take too long...'
        $calArr = Get-OutlookCalendar
    }
    If (!$folderFiles) {
        Write-Output 'Getting Local Files - Errrr, how many files you got?'
        $folderFiles = Get-ChildItem -Path $folderArr -Recurse -File
        $downloadFiles = Get-ChildItem -Path $downloadPath -Recurse -File
    }

    Write-Output "`r`nBuilding Reports...`r`n"
    ForEach ($day in $dateArr) {
        $dateStr = $day.ToString('yyyy-MM-dd')
        Write-Output "TimeScrobbling $dateStr..."
        $outPath = "$outputFld\$dateStr-TimeScrobble.htm"
        $tomorrow = $day.AddDays(1)

        $fileSortProp1 = @{Expression='DirectoryName';Descending = $true}
        $fileSortProp2 = @{Expression='LastWriteTime';Ascending = $true}
        $folderObj = $folderFiles | Where-Object {($_.CreationTime -ge $day -and $_.CreationTime -lt $tomorrow) -or ($_.LastWriteTime -ge $day -and $_.LastWriteTime -lt $tomorrow)} | Select-Object Name, DirectoryName, CreationTime, LastWriteTime | Sort-Object $fileSortProp1,$fileSortProp2
        $downloadObj = $downloadFiles | Where-Object {($_.CreationTime -ge $day -and $_.CreationTime -lt $tomorrow) -or ($_.LastWriteTime -ge $day -and $_.LastWriteTime -lt $tomorrow)} | Select-Object Name, CreationTime, LastWriteTime | Sort-Object lastWriteTime
        $inboxObj = $inboxArr | Where-Object {$_.ReceivedTime -ge $day -and $_.ReceivedTime -lt $tomorrow} | Select-Object ReceivedTime, SenderName, Subject, Importance
        $sentObj = $sentArr | Where-Object {$_.SentOn -ge $day -and $_.SentOn -lt $tomorrow} | Select-Object SentOn, To, Subject, Importance
        $calObj = $calArr | Where-Object {$_.Start -ge $day -and $_.Start -lt $tomorrow} | Select-Object Start, Subject, Duration, Location
    
        If ($slackToken) {
            $slackUsers = Get-SlackUser -Token $slackToken -Presence
            $slackObj = @()
            ForEach ($channel in $slackChannels) {
                $channelMsgs = Get-SlackChannel -Token $slackToken -Name $channel | Get-SlackHistory -Token $slackToken -After $day -Before $tomorrow
                $channelMsgs | ForEach {$_ | Add-Member -MemberType NoteProperty -Name 'Channel' -Value $channel}
                $slackObj += $channelMsgs
            }
            ForEach ($group in $slackGroups) {
                $groupMsgs = Get-SlackGroup -Token $slackToken -Name $group | Get-SlackGroupHistory -Token $slackToken -After $day -Before $tomorrow
                $groupMsgs | ForEach {$_ | Add-Member -MemberType NoteProperty -Name 'Channel' -Value $group}
                $slackObj += $groupMsgs
            }

            $slackFiles = @()
            ForEach ($message in $slackObj) {
                $message.Username = ($slackUsers | Where-Object {$_.ID -eq $message.User} | Select-Object -ExpandProperty Name)
                If ($message.File) {
                    $slackFiles += $message
                }
            }
            If ($slackFiles.count -ne 0) {
                $slackFileObj = @()
                ForEach ($file in $slackFiles) {
                    $SlackFileObj += [PSCustomObject] @{
                        Channel = $file.Channel
                        Timestamp = $file.Timestamp
                        Username = $file.Username
                        Title = $file.File.title
                        Filename = $file.File.name
                        Permalink = $file.File.permalink
                    }
                }
                $slackObj = $slackObj | Where-Object {$slackFiles -notcontains $_}
            }

            $slackSortProp1 = @{Expression='Channel'; Descending=$true}
            $slackSortProp2 = @{Expression='Timestamp'; Ascending=$true}
            $slackObj = $slackObj | Select-Object Channel,Timestamp,Username,Text | Sort-Object $slackSortProp1, $slackSortProp2
        }

        $outBody = @()
        If ($folderObj) {
            $outBody += New-HTMLTable -inputObj $folderObj -headerText 'Personal Files Created/Modified'
        }
        If ($downloadObj) {
            $outBody += New-HTMLTable -inputObj $downloadObj -headerText 'Downloaded Files'
        }
        If ($inboxObj) {
            $outBody += New-HTMLTable -inputObj $inboxObj -headerText 'Received Emails'
        }
        If ($sentObj) {
            $outBody += New-HTMLTable -inputObj $sentObj -headerText 'Sent Emails'
        }
        If ($calObj) {
            $outBody += New-HTMLTable -inputObj $calObj -headerText 'Calendar Entries'
        }
        If ($slackObj) {
            $outBody += New-HTMLTable -inputObj $slackObj -headerText 'Slack Messages'
        }
        If ($slackFileObj) {
            $outBody += New-HTMLTable -inputObj $slackFileObj -headerText 'Slack Files'
        }

        [array]$outBody = "<h1>TimeScrobbler Run for $dateStr</h1>" + $outBody + "<br><h3>Report generated at $((Get-Date).ToString('yyyy-MM-dd'))</h3>"
        $outHTM = ConvertTo-Html -Head $reportHeader -Body $outBody
        $outHTM | Out-File $outPath -Force
    }

    Write-Output "`r`nTimeScrobble complete. Reports available at $outputFld`r`n."
    
    While(!$doneFlag -and !$moreFlag) {
        $doneTest = Read-Host 'TimeScrobble another range? [y/n]'
        switch ($doneTest) {
            'n' {$doneFlag = $true}
            'y' {$moreFlag = $true}
            default {Write-Output "Invalid Entry!`r`n"}
        }
    }
}

Write-Output "`r`nPress any key to exit.."
$x = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')