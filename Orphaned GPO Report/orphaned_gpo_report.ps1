<#

.SYNOPSIS 
    This script finds all orphaned GPOs 

.DESCRIPTION
    This script fetches all orphaned GPOs and then outputs 	the results to a CSV file and also to an email which 
	is then sent to a specificed list.
	
	Before use edit the script if you want the results sent via email.

.RELATED LINKS
    https://github.com/CptSternn/Powershell

.NOTES
    Version:      1.0
    
	Release Date: 11-12-2019
	Last Modified: 11-12-2019
   
    Author:	Wesley Whitworth

.EXAMPLE
    orphaned_gpo_report.ps1
#>

# Configure Email
$SendEmail = "FALSE" # TRUE or FALSE
$ToAddress = 'Username <user@domain.com>'
$FromAddress = 'Reports <reports@domain.com>'
$SMTPserver = 'SMTP-Server'

# Setup variables
$startDTM = (Get-Date)
$Domain = Get-ADDomain -Current LocalComputer | Select DNSroot
$DomainName = $Domain.DNSroot
$UserCounter = 0
$Subject = "Orphaned GPO Report ($DomainName)" + " - " + (Get-Date -format d)
$OutFile = ".\OrphanedGPOs - " + (Get-Date).tostring("dd-MM-yyyy-hh-mm-ss") + ".csv"

# Write headers to CSV
Add-Content $OutFile "Display Name, Owner, Created, Modified"

# Fetch empty GPO list
Write-Host "`nFetching empty GPOs..." -ForegroundColor Cyan
$GPOs = Get-GPO -All | Sort-Object displayname | Where-Object { If ( $_ | Get-GPOReport -ReportType XML | Select-String -NotMatch "<LinksTo>" ) {$_.DisplayName } } | select DisplayName, Owner, CreationTime, ModificationTime

# Fetch GPO details
ForEach ($GPO in $GPOs) {

	$GPOCounter++
	$DisplayName = $GPO.DisplayName
	$Owner = $GPO.Owner
	$Created = $GPO.CreationTime
	$Modified = $GPO.ModificationTime
	
	# Update progress bar
	Write-Progress -activity "Processing $DisplayName" -status "$GPOCounter Out Of $($GPOs.Count) completed" -percentcomplete ($GPOCounter / $GPOs.Count*100) 
	
	# Add GPO details to output
	Add-Content $OutFile "$DisplayName,$Owner,$Created,$Modified"
	$t_file_content = $t_file_content + "<tr class=greyback><td>$DisplayName</td><td>$Owner</td><td>$Created</td><td>$Modified</td>"
	Write-Host "$DisplayName"
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
$t_file_header = "<div class=reporttitle>Inactive Computer Report ($DomainName)</div><table><tr class=blueback><td>Display Name</td><td>Enabled</td><td>Last Logon</td><td>Distinguished Name</td></tr>"
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
$infobox = $infobox + "<b>Description:</b> This script finds all empty GPOs in a domain<br>"
$infobox = $infobox + "<b>Script run time:</b> $ts_display<br>"
$infobox = $infobox + "<br>" 
$infobox = $infobox + "</td></tr></table>"

$html_header = "<html><head>" + $a + "</head><body>"
$body_main = $body_main + $t_file + $key + $infobox
$html_footer = "</body></html>"

$body = $html_header + $body_main + $html_footer

If ($SendEmail -eq "TRUE") {
	# Send the email
	Send-MailMessage -To $ToAddress -From $FromAddress -Subject $Subject -SMTPserver $SMTPserver -Body $body -BodyAsHtml
}

# Output details to screen
$totalscripttime = (($endDTM - $startDTM).totalminutes)
$totalscripttime = [System.Math]::Round($totalscripttime, 0)
Write-Host "`nTotal Empty GPOs Found: $GPOCounter" -ForegroundColor Cyan
Write-Host "Script run time: $totalscripttime minutes`n" -ForegroundColor Cyan