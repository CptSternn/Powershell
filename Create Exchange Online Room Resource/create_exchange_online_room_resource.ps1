<#

.SYNOPSIS 
    This script creates a Room Resources in Exchange Online for use with Teams

.DESCRIPTION
    This script will first check to see if the Exchange Online and Azure AD modules are installed and if not
	it will install them. Next it will get the users Azure admin account details and then log the user into 
	Exchange and AzureAD (With MFA prompt). Once connected it will read data from a CSV file found in the same
	folder as the script and use the data to create new Room Resources in Exchange Online and set the needed
	permissions as required for Teams in AzureAD. This process is outlined in this article:
	
	https://learn.microsoft.com/en-us/microsoftteams/rooms/with-office-365
	
	Please note that a license for the new Room Resource *must* be assigned manually.	
	Before use edit the CSV file and enter the room data.

.RELATED LINKS
    https://github.com/CptSternn/Powershell

.NOTES
    Version:      1.0
    
	Release Date: 17-02-2023
	Last Modified: 17-02-2023
   
    Author:	Wesley Whitworth

.EXAMPLE
    create_exchange_online_room_resource.ps1
#>
# Setup variables
$startDTM = (Get-Date)
$filename = '.\rooms.csv'

# Get rooms list from CSV file
$Rooms = Import-CSV $filename
$TotalRooms = $rooms.count + 1

Write-Host "$TotalRooms Rooms found in CSV file `n" -Foreground Green

# Check for the two modules needed and install if not found
Write-Host "Checking for Exchange Online and Azure AD modules...`n" -Foreground Cyan

If (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
    Write-Host "Module ExchangeOnlineManagement found" -ForegroundColor Green
	Import-Module ExchangeOnlineManagement
} 
Else {
    Write-Host "Installing ExchangeOnlineManagement Module" -ForegroundColor Yellow
	Install-Module ExchangeOnlineManagement
	Import-Module ExchangeOnlineManagement
}

If (Get-Module -ListAvailable -Name AzureAD) {
    Write-Host "Module AzureAD found `n" -ForegroundColor Green
	Import-Module AzureAD
} 
Else {
    Write-Host "Installing AzureAD Module `n" -ForegroundColor Yellow
	Install-Module AzureAD
	Import-Module AzureAD
}

# Get users admin account name
Add-Type -AssemblyName Microsoft.VisualBasic
$User = [Microsoft.VisualBasic.Interaction]::InputBox('Azure Admin Username (user@domain): ', 'Azure Admin Login Credentials Prompt', "Enter Azure Admin UPN here")

Write-Host "`nAzure Administrator Account: $user `n" -ForegroundColor Cyan
Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow

# Logon to Exchange Online
Try {
	Connect-ExchangeOnline -UserPrincipalName $user -ShowBanner:$false
}
Catch {
	$_.Exception.Message ; Exit 1
}

Write-Host "Connection to Exchange Online successful. `n" -ForegroundColor Green

Write-Host "Connecting to AzureAD..." -ForegroundColor Yellow

# Logon to AzureAD
Try {
	$AADstatus = Connect-AzureAD -AccountID $User
}
Catch {
	$_.Exception.Message ; Exit 1
}

Write-Host "Connection to Azure successful. `n" -ForegroundColor Green

# Grab the room data from the CSV and add it to Exchange Online

Foreach ($Room in $Rooms) {

	$RoomCounter++
	$RoomName = $Room.Displayname
	$RoomEmail = $Room.EmailAddress
	$RoomAlias = $Room.ExchangeAlias
	$RoomPassword = $Room.Password
	
	# Create the Mailbox
	Try {
		New-Mailbox -MicrosoftOnlineServicesID $RoomEmail -Name $RoomName -Alias $RoomAlias -Room -EnableRoomMailboxAccount $true  -RoomMailboxPassword (ConvertTo-SecureString -String $RoomPassword -AsPlainText -Force)
	}
	Catch {
		$_.Exception.Message ; Exit 1
	}
	
	Write-Host "`n$RoomEmail : Mailbox creation successful. `n" -ForegroundColor Green
	
	# Set Calendar Processing for the new mailbox
	Try {
		Set-CalendarProcessing -Identity $RoomAlias -AutomateProcessing AutoAccept -AddOrganizerToSubject $false -DeleteComments $false -DeleteSubject $false -ProcessExternalMeetingMessages $true -RemovePrivateProperty $false -AddAdditionalResponse $true -AdditionalResponse "This is a Microsoft Teams Meeting room!"
	}
	Catch {
		$_.Exception.Message ; Exit 1
	}

	Write-Host "$RoomName : Calendar Processing configuration successful. `n" -ForegroundColor Green
	Write-Host "Pausing for 20 seconds while the new mailbox syncs before applying Password Expiration policies. `n" -ForegroundColor Cyan
	Start-Sleep -Seconds 20
	
	# Disable password expiration on the Room Resource account
	Try {
		Set-AzureADuser -ObjectID $RoomEmail -PasswordPolicies DisablePasswordExpiration
	}
	Catch {
		$_.Exception.Message ; Exit 1
	}
	
	Write-Host "Room Created: $RoomName / $RoomEmail / $RoomAlias " -ForegroundColor Green
}

# Disconnect the remote admin sessions
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-AzureAD

# Output the final results
$endDTM = (Get-Date)
$totalscripttime = (($endDTM - $startDTM).totalminutes)
$totalscripttime = [System.Math]::Round($totalscripttime, 0)
Write-Host "`nExchange Online Room Resources Created: $RoomCounter`nScript run time: $totalscripttime minutes`n" -ForeGround Cyan
