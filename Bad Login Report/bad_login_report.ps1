<#

.SYNOPSIS 
    This script reads the security event log from multiple servers and outputs the data to a HTML formated 
    email. 

.DESCRIPTION
    This script reads the security event log from multiple servers and checks the data, formats it and then
    outputs the data to a CSV file and if enabled a HTML formated email. 

.RELATED LINKS
    https://github.com/CptSternn/Powershell

.NOTES
    Version:      1.0
    
    Release Date: 12-05-2019
	Last Modified: 20-02-2023
   
    Author:	Wesley Whitworth

.EXAMPLE
    bad_login_report.ps1
#>
# Configure Email
# *Don't enable for searches larger than 24 hours as it will cause Outlook viewing issues with the large results!
$SendEmail = "FALSE" # TRUE or FALSE
$ToAddress = 'Username <user@domain.com>'
$FromAddress = 'Reports <reports@domain.com>'
$SMTPserver = 'SMTP-Server'

# Set time period to check
#$StartTime = (get-date).AddDays(-4) # Get last 4 days
#$StartTime = (get-date).AddDays(-1) # Get last 24 hours
#$StartTime = (get-date).AddMinutes(-30) # Get last 30 minutes
$StartTime = (get-date).AddHours(-1) # Get last hour

# Setup variables
$startDTM = (Get-Date)
$Subject = "Bad Logon Report" + " - " + (Get-Date -format d)
$OutFile = ".\BadLogins - " + (Get-Date).tostring("dd-MM-yyyy-hh-mm-ss") + ".csv"
$Servers = Get-ADDomainController -filter * | Select-Object name | Sort name
$EventId = 4771, 4740

$LogFilter = @{
    Logname   = 'Security'
    ID        = $EventID
    StartTime = $StartTime
}

Write-Host "Fetching Event Logs..." -Foreground Cyan

ForEach ($Server in $Servers) {
	
	$ServerCounter++
	$EventCounter = $NULL
	
	$Server = $server.name
	Write-Host "`nChecking $Server..." -Foreground Cyan

	# Update progress bar
	Write-Progress -activity "Processing $Server" -status "$ServerCounter Out Of $($Servers.Count) completed" -percentcomplete ($ServerCounter / $servers.Count*100) 


    Try {
        $events = Get-WinEvent -FilterHashtable $LogFilter -ComputerName $Server -ErrorAction SilentlyContinue
    }
    Catch {
        $_.Exception.Message
    }
	
	If ($Events) {
			
		ForEach ($event in $events) {
				
			$hostname = $null
		
			$eventXML = [xml]$Event.ToXml() 
		
			ForEach ($attr in $eventXML.Event.EventData.Data) { 
				If ($attr.name -eq "TargetUserName") {
					$Username = $attr.'#text'
				}
				If ($attr.name -eq "Status") {
					$Status = $attr.'#text'
				}
				If ($attr.name -eq "IPaddress") {
					$IPaddress = $attr.'#text'
				}
			}
		
			$Eid = $Event.Id
			$TimeCreated = $Event.TimeCreated
		
			Switch ($Status) {
				"0x18" { $StatusDesc = "Bad Password" }
				"0x17" { $StatusDesc = "Password Expired" }
				"0x12" { $StatusDesc = "Account Locked Out" }
				"0x25" { $StatusDesc = "Clock Too Far Off To Sync" }
				Default { $StatusDesc = "Unknown Error Code" }
			}

			$IPaddress = $IPaddress.Replace("::ffff:", "")
		
			$Username = $Username.trim().ToLower()

			# Do a reverse dns lookup for the hostname
		
			Try {
				$hostname = [System.Net.Dns]::GetHostByAddress($IPaddress).HostName 
			}
			Catch {
				$hostname = "NO_DNS_ENTRY"
			}
		
			If ($Username -NotLike '*$') {
				Add-Content $OutFile "$TimeCreated,$Server,$Username,$Hostname,$IPaddress,$status,$statusdesc"
				Write-Host "$TimeCreated $Server $Username $Hostname $IPaddress $status $statusdesc"
				$t_file_content = $t_file_content + "<tr class=greyback><td>$TimeCreated</td><td>$Server</td><td>$EID</td><td style='background-color:$tabcol'>$Username</td><td>$Hostname</td><td>$IPaddress</td><td>$Status</td><td>$StatusDesc</td>"
				$EventCounter++
			}
			
			# Update second progress bar
			Write-Progress -activity "Processing $EID at $TimeCreated" -status "$EventCounter Out Of $($events.Count) completed" -percentcomplete ($EventCounter / $events.Count*100) -id 1
		}
	}
	Else {
		Write-Host "No matching events found"
	}
}

# Define email CSS 
$a = "<style>"
$a = $a + "BODY{background-color:white;}"
$a = $a + "TH.norm{border-width: 1px;padding:3px;border-style:solid;border-color:#cdcdcd;background-color:#efefef;font-family:verdana,arial;font-size:11px}"
$a = $a + "TD.norm{border-width: 1px;padding:3px;border-style:solid;border-color:#cdcdcd;background-color:#efefef;font-family:verdana,arial;font-size:11px}"
$a = $a + "TR.norm{border-width: 1px;padding:3px;border-style:solid;border-color:#cdcdcd;background-color:#efefef;font-family:verdana,arial;font-size:11px}"
$a = $a + "tr.blueback td{border-width: 1px;padding:3px;border-style:solid;border-color:#cdcdcd;background-color:#cedeef;font-family:verdana,arial;font-size:11px}"
$a = $a + "tr.bluebacktwo td{text-align:center;border-width: 1px;padding:3px;border-style:solid;border-color:#cdcdcd;background-color:#cedeef;font-family:verdana,arial;font-size:11px;color:#ffffff;}"
$a = $a + "tr.greyback td{border-width: 1px;padding:3px;border-style:solid;border-color:#cdcdcd;background-color:#efefef;font-family:verdana,arial;font-size:11px;}"
$a = $a + "tr.scale td{height:100px;width:30px;text-align:center;background-color:#efefef;vertical-align:top;border-top-style:solid;border-top-width:1px;}"
$a = $a + ".reporttitle {color:#5CADFF;font-family:Calibri,verdana,arial;font-size:28px;}"
$a = $a + "</style>"

# Create the email
$t_file_header = "<div class=reporttitle>Failed Logon Events Report</div><table><tr class=blueback><td>Date/Time</td><td>Server</td><td>EventID</td><td>Username</td><td>Hostname</td><td>IP</td><td>ID</td><td>Status</td></tr>"
$t_file_footer = "</table>"
$t_file = $t_file_header + $t_file_content + $t_file_footer

$endDTM = (Get-Date)
$ts = ($endDTM - $startDTM)
$ts_display = '{0:00}h {1:00}m {2:00}s' -f $ts.Hours, $ts.Minutes, $ts.Seconds

$ScriptPath = $MyInvocation.MyCommand.Path
$host_name = $env:COMPUTERNAME

$infobox = "<br><br><table style='width:100%'><tr class=greyback><td><b>Script:</b> $ScriptPath<br>"
$infobox = $infobox + "<b>Server:</b> $host_name<br>"
$infobox = $infobox + "<b>Scheduled Task:</b> Not set<br>"
$infobox = $infobox + "<b>Description:</b> This script checks failed logon events for a specific date range<br>"
$infobox = $infobox + "<b>Script run time:</b> $ts_display<br>"
$infobox = $infobox + "<br>" 
$infobox = $infobox + "</td></tr></table>"

$html_header = "<html><head>" + $a + "</head><body>"
$body_main = $body_main + $t_file + $key + $infobox
$html_footer = "</body></html>"

$body = $html_header + $body_main + $html_footer


# Send the email
If ($SendEmail -eq "TRUE") {
	# Send the email
	Send-MailMessage -To $ToAddress -From $FromAddress -Subject $Subject -SMTPserver $SMTPserver -Body $body -BodyAsHtml
}

# Output details to screen
$totalscripttime = (($endDTM - $startDTM).totalminutes)
$totalscripttime = [System.Math]::Round($totalscripttime, 0)
Write-Host "`nTotal Servers Checked: $ServerCounter" -ForegroundColor Cyan
Write-Host "Script run time: $totalscripttime minutes`n" -ForegroundColor Cyan