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
    $outputArr += ("<br><span style=`'font-weight: bold;`'>",$headerText,'</span>' -join '')
    $outputArr += $inputObj | ConvertTo-HTML -Fragment | Set-AlternatingRows -CSSOddClass odd -CSSEvenClass even
    return $outputArr
}

Function Set-AlternatingRows {
  [CmdletBinding()]
   	Param(
       	[Parameter(Mandatory,ValueFromPipeline)]
        [string]$Line,
       
   	    [Parameter(Mandatory)]
       	[string]$CSSEvenClass,
       
        [Parameter(Mandatory)]
   	    [string]$CSSOddClass
   	)
	Begin {
		$ClassName = $CSSEvenClass
	}
	Process {
		If ($Line.Contains('<tr><td>'))
		{	$Line = $Line.Replace('<tr>',"<tr class=""$ClassName"">")
			If ($ClassName -eq $CSSEvenClass)
			{	$ClassName = $CSSOddClass
			}
			Else
			{	$ClassName = $CSSEvenClass
			}
		}
		Return $Line
	}
}

Function Set-Folder {
# Function to check if a folder exists, and create it if not.
	param($path)
			
	If (!(Test-Path "filesystem::$path" -ErrorAction SilentlyContinue)) {
		New-Item $path -Type Directory -Force | Out-Null
	}
}

$reportHeader = @"
<style>
  body {
    font-family: "Arial";
    font-size: 10pt;
    color: #4C607B;
    }
  th, td { 
    border: 1px solid #e57300;
    border-collapse: collapse;
    padding: 5px;
    }
  th {
    font-size: 1.2em;
    text-align: left;
    background-color: #003366;
    color: #ffffff;
    }
  td {
    color: #000000;
    }
  .even { background-color: #ffffff; }
  .odd { background-color: #bfbfbf; }
</style>
"@


Write-Output 'TimeScrobbler v1.0 - kittiah@gmail.com'
# Work out path, import Slack module (TEMP: Until the Groups functionality is added to master branch)
$scriptPath = Split-Path $MyInvocation.MyCommand.Path -Parent

# Read in the User.conf and turn into live variables
$importConfig = Get-Content $scriptPath\User.conf | Where-Object {($_ -notlike '#*') -and ($_)}
$importConfig | ForEach {
    $splitVar = $_.Split('=')
    If ($splitVar[1] -like '*,*') {
        [array]$value = $splitVar[1].Trim().Split(',')
    }
    Else {
        $value = $splitvar[1].Trim()
    }
    Set-Variable -Scope Script -Name $splitVar[0].Trim() -Value $value
}

# Import PSSlack if a token has been provided
If ($personalSlackKey -ne '') {
    Import-Module $scriptPath\PSSlack -Force | Out-Null
}

# Create the output folder if it doesn't already exist
Set-Folder $outputFld

# Work out the days we need to generate reports for
$startDate = Get-Date -Date $startDay
$endDate = Get-Date -Date $endDay
$difference = New-TimeSpan -Start $startdate -End $enddate
$days = [Math]::Ceiling($difference.TotalDays)+1
$dateArr = @()
1..$days | ForEach-Object {
  $dateArr += $startdate
  $startdate = $startdate.AddDays(1)
}

$folderArr += [Environment]::GetFolderPath('Desktop')
$folderArr += [Environment]::GetFolderPath('Desktop')
$downloadPath = Get-ItemProperty 'Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' | Select-Object -ExpandProperty '{374DE290-123F-4565-9164-39C4925E467B}'

# Build out the data sources for the reports
# Please note, these bits take fucking ages. Go make a sandwich or three.
Write-Output 'Getting Outlook Inbox - This may take some time... No, seriously. Make a sandwich.'
$inboxArr = Get-OutlookInbox
Write-Output 'Getting Outlook Sent Items - This could be a while...'
$sentArr = Get-OutlookSent
Write-Output 'Getting Outlook Calendar - Hopefully wont take too long...'
$calArr = Get-OutlookCalendar

Write-Output 'Getting Local Files'
$folderFiles = Get-ChildItem -Path $folderArr -Recurse -File
$downloadFiles = Get-ChildItem -Path $downloadPath -Recurse -File

Write-Output 'Building Reports...'
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
    
    If ($personalSlackKey) {
    # Please note, requires customised version of PSSlack with Group support to function properly @ 17/08
    $slackUsers = Get-SlackUser -Token $personalSlackKey -Presence
    $slackObj = @()
        ForEach ($channel in $slackChannels) {
            $channelMsgs = Get-SlackChannel -Token $personalSlackKey -Name $channel | Get-SlackHistory -Token $personalSlackKey -After $day -Before $tomorrow
            $channelMsgs | ForEach {$_ | Add-Member -MemberType NoteProperty -Name 'Channel' -Value $channel}
            $slackObj += $channelMsgs
        }
        ForEach ($group in $slackGroups) {
            $groupMsgs = Get-SlackGroup -Token $personalSlackKey -Name $group | Get-SlackGroupHistory -Token $personalSlackKey -After $day -Before $tomorrow
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

    [array]$outBody = "<h1>TimeScrobbler Run for $dateStr</h1>" + $outBody + "<br><h3>Report generated at $(Get-Date)</h3>"
    $outHTM = ConvertTo-Html -Head $reportHeader -Body $outBody
    $outHTM | Out-File $outPath -Force
}

Write-Output "`r`nTimeScrobble complete. Reports available at $outputFld`r`nPress any key to exit."
$x = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')