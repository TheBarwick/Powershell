<#
Company: CDW
Name: MailboxDetailedDiscovery.ps1
Author: Matthew Sherwood
Author Email: matt.sherwood@s3.cdw.com
Date Created: 4/20/2016

***Requirements***
	Script must be run on system that meet the following requirements
	-List of UPNs in a CSV file
		+ List cannot contain a header. Only a list of UPNs.
	-Microsoft Online Services Sign-In Assistant for IT Professionals RTW
		+ http://go.microsoft.com/fwlink/?LinkID=286152
	-Azure Active Directory Module for Windows PowerShell (64-bit version)
		+ http://go.microsoft.com/fwlink/p/?linkid=236297
	-Server Remote Managment Tools - Active Directory Management Tools
		+ https://technet.microsoft.com/en-us/library/cc816817(v=ws.10).aspx

***Change Log***
5/5/16
	- Removed the "AD Account Expires" human redable output. Now the script only shows an expiration date if the account expires. 

#>


#Adjust these variables to fit your needs

	#Import file. Specify CSV file that only includes list of UPNs, without a header. 
	$ImportFile = ".\UserImport.csv"

	#Export File. Specify the directory and file name you would like to save your report in. 
	#Default saves report to same directory as script, with date/time info.
	$DateTime = "$((Get-Date).ToString('MM-dd-yyyy_hh-mm-ss'))"
	$ExportFile = ".\$DateTime MailboxReport.csv"
	#$ExportFile = Read-Host "Enter the location to save the CSV report (Eg. C:\Scripts\MailboxReport.csv)"

	#To verify that the accounts are AD Synced, enter the ProxyAddress search string you would like to use (example "*.onmicrosoft.com")
	$ProxyAddressSearch = "*mail.onmicrosoft.com"
	
	#Enter the number of maximum number (in days) to limit the report to only show devices that have attempted to sync within this number of of days.  
	$MobileDeviceAge = 30					# Compared with ActiveSyncDeviceStatistics.LastSyncAttemptTime





#------------------------- Do Not Change Below This Line -------------------------#
#------------------------- Do Not Change Below This Line -------------------------#
#------------------------- Do Not Change Below This Line -------------------------#


#Import modules
Import-Module MSOnline
Import-Module ActiveDirectory

#Create null array
$MailboxData = @()
$MailboxDataOutput = @()

### Connect to services ###
#Check for Local Exchnage Connection.  If fails make a connection.
Try {$test = Get-PSSession -Name "ON-Prem-Exchange" -ErrorAction stop}
Catch{Invoke-Expression .\ConnectLocalExchange.ps1 }
Finally{ }

#Connect to Azure AD
Try {$test = Get-MsolDomain -ErrorAction stop | where { ($_.name -like "*.mail.onmicrosoft.com") } }
Catch{Invoke-Expression .\Connect365.ps1 }
Finally{ }

cls

#Setup Shell
$bg = (Get-Host).UI.RawUI
$BG.BackgroundColor = "Black"
cls

#Set date variable
$now = Get-Date


#Import user list
$Users = Get-Content $ImportFile
#$Users = Get-mailbox -resultsize unlimited | where {$_.RecipientTypeDetails -ne "DiscoveryMailbox"}


Write-Host "
Starting to gather data for each user...
"

Foreach ($User in $Users) {
	#Create variables and clear previous variables.
	#General Info
	$MBXUser = @()
	$MBXStatistics = @()
	$MBXTotalItemSize = @()
	$Recipient = @()
	$ADUser = @()
	$ActiveSyncDeviceStatistics = @()
	$IsCloudSynced = @()
	$OnMicrosoftEmailAddress = @()
	$MailboxIsMigrated = @()
	$MailboxLocation = @()
	$targetAddress = @()
	$ForwardingAddress = @()
	#Mobile Device Info
	$HasMobileDevice = @()
	$MobileDeviceIndex = 0
	$MobileDeviceFriendlyName = @()
	$MobileDeviceModel = @()
	$MobileDeviceOS = @()
	$MobileDeviceLastSuccessSync = @()
	$MobileDeviceFriendlyNameString = ""
	$MobileDeviceModelString = ""
	$MobileDeviceOSString = ""
	$MobileDeviceLastSuccessSyncString = ""
	#Permissions Info
	$SendAs = @()
	$SendOnBehalf = @()
	$SendAsString = @()
	$FullAccessNotInherited = @()
	$FullAccessNotInheritedString = @()
	
	#Verify that user exists in AD, and also store AD properties in the ADUser variable
	$ADUser = Get-ADUser -Filter {UserPrincipalName -like $User} -Properties * -ErrorAction SilentlyContinue
	if ($ADUser -eq $null) {
		Write-Host "$User doesn't exist in Active Directory. Skipping user." -ForegroundColor Red
	} else {
	Write-Host "Gathering data about $User" -ForegroundColor Green
		
	#Create the MailboxData object
	$MailboxData = New-Object PSObject

	#Gather data and store in these variables
	$MBXUser = Get-OnPremMailbox -Identity $User -ErrorAction SilentlyContinue
	$MBXStatistics = Get-OnPremMailboxStatistics -Identity $User -ErrorAction SilentlyContinue
	$Recipient = Get-OnPremRecipient -Identity $User -ErrorAction SilentlyContinue
	$ActiveSyncDeviceStatistics = Get-OnPremActiveSyncDeviceStatistics -Mailbox $User -WarningAction SilentlyContinue -ErrorAction SilentlyContinue  
	$SendAs = Get-OnPremADPermission $ADUser.DistinguishedName -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | `
		Where-Object {$_.ExtendedRights -like "Send-As" `
		-and $_.User -notlike "NT AUTHORITY\SELF" `
		-and $_.User -notlike "SIC\Exchange Domain Servers" `
		-and $_.User -notlike "SIC\svc_BESAdmin" `
		-and $_.User -notlike "STANDARD\Exchange Domain Servers" `
		-and $_.User -notlike "STANDARD\Exchange Services" `
		-and $_.User -notlike "SIC\ExngAdmin" `
		-and $_.User -notlike "S-1*" `
		-and !$_.IsInherited} | ForEach-Object {$_.User}
	$SendOnBehalf = [string]$MBXUser.GrantSendOnBehalfTo
	$FullAccessNotInherited = Get-OnPremMailboxPermission $ADUser.CanonicalName -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | `
		Where-Object {$_.AccessRights -eq "FullAccess" `
		-and $_.User -notlike "S-1*" `
		-and !$_.IsInherited} | ForEach-Object {$_.User}
	#$FullAccess = Get-MailboxPermission $ADUser.CanonicalName -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Where-Object {$_.AccessRights -eq "FullAccess"} | ForEach-Object {$_.User}
	$SendAsString = [string]$SendAs
	$FullAccessNotInheritedString = [string]$FullAccessNotInherited
	#$ActiveSyncDevice = Get-ActiveSyncDevice -Mailbox $MBXUser.Identity
	
	#Convert the mailbox total item size to standard MB output...
	$MBXTotalItemSize = Get-OnPremMailboxStatistics -Identity $User -ErrorAction SilentlyContinue | Select @{name="TotalItemSize"; expression={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}}

	#Check to see if account is synced to Azure AD
	$test = Get-MsolUser -UserPrincipalName $User -ErrorVariable errorVariable -ErrorAction SilentlyContinue
	if ($errorVariable -ne $null) {
		$IsCloudSynced = "False"
	} else {
		$IsCloudSynced = "True"
	}

	#Get OnMicrosoft.EmailAddress from AD
	if ($IsCloudSynced -eq "True") {
		$ADUser.proxyAddresses | `
		foreach {
			if($_ -like $ProxyAddressSearch){
				$split = $_ -split ":"
				$OnMicrosoftEmailAddress = $split[1]
			}
		}
	} else {
		$OnMicrosoftEmailAddress = ""
	}

	#See where mailbox is located
	if ($Recipient.RecipientTypeDetails -eq "UserMailbox") {
		#User's mailbox is still on prem
		$MailboxIsMigrated = "False"
		$MailboxLocation = "On-Prem"
		$targetAddress = ""
		$ForwardingAddress = $MBXUser.ForwardingAddress  
	} elseif ($Recipient.RecipientTypeDetails -eq "RemoteUserMailbox") {
		#User's mailbox is still online
		$MailboxIsMigrated = "True"
		$MailboxLocation = "Online"
		$targetAddress = $ADUser.targetAddress
		$ForwardingAddress = ""
	} elseif ($Recipient.RecipientTypeDetails -eq "SharedMailbox") {
		#User's mailbox is shared
		$MailboxIsMigrated = "False"
		$MailboxLocation = "On-Prem"
		$targetAddress = $ADUser.targetAddress
		$ForwardingAddress = ""
	} else {
		$MailboxIsMigrated = ""
		$MailboxLocation = ""
		$targetAddress = ""
		$ForwardingAddress = ""
	}

	#Add first set of Mailbox data properties to the MailboxDataOutput object
	$MailboxData | Add-Member NoteProperty -Name "Display Name" -Value $ADUser.displayname
	$MailboxData | Add-Member NoteProperty -Name "SamAccountName" -Value $ADUser.SamAccountName
	$MailboxData | Add-Member NoteProperty -Name "First Name" -Value $ADUser.GivenName
	$MailboxData | Add-Member NoteProperty -Name "Last Name" -Value $ADUser.Surname
	$MailboxData | Add-Member NoteProperty -Name "UPN" -Value $ADUser.UserPrincipalName
	$MailboxData | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Recipient.PrimarySmtpAddress
	$MailboxData | Add-Member NoteProperty -Name "Alias" -Value $Recipient.Alias
	$MailboxData | Add-Member NoteProperty -Name "Account Synced to Azure AD" -Value $IsCloudSynced
	$MailboxData | Add-Member NoteProperty -Name "mail.onmicrosoft.com" -Value $OnMicrosoftEmailAddress
	$MailboxData | Add-Member NoteProperty -Name "Mailbox Type" -Value $Recipient.RecipientTypeDetails
	$MailboxData | Add-Member NoteProperty -Name "ADEnabled" -Value $ADUser.Enabled
	$MailboxData | Add-Member NoteProperty -Name "AD Account Expiration Date" -Value $ADUser.AccountExpirationDate
	$MailboxData | Add-Member NoteProperty -Name "Last Logged in Mailbox" -Value $MBXStatistics.LastLogonTime
	$MailboxData | Add-Member NoteProperty -Name "Department" -Value $ADUser.Department
	$MailboxData | Add-Member NoteProperty -Name "Title" -Value $ADUser.Title
	$MailboxData | Add-Member NoteProperty -Name "targetAddress" -Value $targetAddress
	$MailboxData | Add-Member NoteProperty -Name "UM Enabled" -Value $Recipient.UMEnabled
	$MailboxData | Add-Member NoteProperty -Name "UM Extension" -Value $MBXUser.Extensions
	$MailboxData | Add-Member NoteProperty -Name "UM Mailbox Policy" -Value $Recipient.UMMailboxPolicy
	$MailboxData | Add-Member NoteProperty -Name "Managed Folder Policy" -Value $Recipient.ManagedFolderMailboxPolicy
	$MailboxData | Add-Member NoteProperty -Name "Email Address Policy Enabled" -Value $Recipient.EmailAddressPolicyEnabled
	$MailboxData | Add-Member NoteProperty -Name "Hidden From Address Lists Enabled" -Value $Recipient.HiddenFromAddressListsEnabled
	$MailboxData | Add-Member NoteProperty -Name "Mobile Phone Number" -Value $ADUser.MobilePhone
	$MailboxData | Add-Member NoteProperty -Name "AD Phonenumber" -Value $ADUser.telephoneNumber
	$MailboxData | Add-Member NoteProperty -Name "Total Item Size (MB)" -Value $MBXTotalItemSize.TotalItemSize
	$MailboxData | Add-Member NoteProperty -Name "Item Count" -Value $MBXStatistics.ItemCount
	$MailboxData | Add-Member NoteProperty -Name "Mailbox Already Migrated" -Value $MailboxIsMigrated
	$MailboxData | Add-Member NoteProperty -Name "Skype Phone Number" -Value $ADUser.'msRTCSIP-Line'
	$MailboxData | Add-Member NoteProperty -Name "SIP Address" -Value $ADUser.'msRTCSIP-PrimaryUserAddress'

	#If user has an Active Sync Device, gather this data. If not, null all data.
	if ($ActiveSyncDeviceStatistics -eq $null) {
		$HasMobileDevice = "False"
		$MobileDeviceFriendlyName = ""
		$MobileDeviceModel = ""
		$MobileDeviceOS = ""
		$MobileDeviceLastSuccessSync = ""
	} else {
		$HasMobileDevice = "True"
		#Verify that the device has synced within the specified number of days. If it has, gather the following data about each device and add it to the report.
		Foreach ($MobileDevice in $ActiveSyncDeviceStatistics){
       		$lastsyncattempt = $MobileDevice.LastSyncAttemptTime
					$MobileDeviceSyncAge = ($now - $lastsyncattempt).Days
			if ($lastsyncattempt -eq $null) {																				# If device hasn't attempted to sync ever...
				$MobileDeviceSyncAge = "Never"
			} elseif ($MobileDeviceSyncAge -le $MobileDeviceAge) {									# If device attempted to sync within the specified number of days, add to report.
				Write-Host "Device $($MobileDevice.DeviceFriendlyName) has synced within $MobileDeviceAge days" -ForegroundColor DarkGreen
				Write-Host "	Adding mobile device to report" -ForegroundColor DarkGreen
				Write-Host "	Last attempted sync was $MobileDeviceSyncAge days ago" -ForegroundColor DarkGreen
				Write-Host "	Device GUID is $($MobileDevice.GUID)" -ForegroundColor DarkGreen
				$MobileDeviceIndex ++																									# Adds 1 to index each time a valid device is processed
				$MobileDeviceFriendlyName += "$MobileDeviceIndex) $($MobileDevice.DeviceFriendlyName)"
				$MobileDeviceModel += "$MobileDeviceIndex) $($MobileDevice.DeviceModel)"
				$MobileDeviceOS += "$MobileDeviceIndex) $($MobileDevice.DeviceOS)"
				$MobileDeviceLastSuccessSync += "$MobileDeviceIndex) $($MobileDevice.LastSuccessSync)"

				#Convert values to string for CSV report
				$MobileDeviceFriendlyNameString = [string]$MobileDeviceFriendlyName
				$MobileDeviceModelString = [string]$MobileDeviceModel
				$MobileDeviceOSString = [string]$MobileDeviceOS
				$MobileDeviceLastSuccessSyncString = [string]$MobileDeviceLastSuccessSync

			} elseif ($MobileDeviceSyncAge -gt $MobileDeviceAge) {							# If device attempted to sync over the specified number of days, don't add to report.
				Write-Host "Device $($MobileDevice.DeviceFriendlyName) hasn't synced within $MobileDeviceAge days" -ForegroundColor DarkMagenta
				Write-Host "	Not including in report" -ForegroundColor DarkMagenta
				Write-Host "	Last attempted sync was $MobileDeviceSyncAge days ago" -ForegroundColor DarkMagenta
				Write-Host "	Device GUID is $($MobileDevice.GUID)" -ForegroundColor DarkMagenta
			} else {																														# There was a script error, and it needs to be fixed...
				Write-Host "Device $($MobileDevice.DeviceFriendlyName) didn't meet elseif conditions. Script needs to be fixed."
				Write-Host "	Device GUID is $($MobileDevice.GUID)" -ForegroundColor Red
			} 
		}
	}
	
	#Add second set of Mailbox data properties to the MailboxDataOutput object
	$MailboxData | Add-Member NoteProperty -Name "Has an ActiveSync Device" -Value $HasMobileDevice
	$MailboxData | Add-Member NoteProperty -Name "Mobile Device Count" -Value $MobileDeviceIndex
	$MailboxData | Add-Member NoteProperty -Name "Friendly Name" -Value $MobileDeviceFriendlyNameString
	$MailboxData | Add-Member NoteProperty -Name "Last Successful Sync" -Value $MobileDeviceLastSuccessSyncString
	$MailboxData | Add-Member NoteProperty -Name "Device Model" -Value $MobileDeviceModelString
	$MailboxData | Add-Member NoteProperty -Name "OS version" -Value $MobileDeviceOSString
	$MailboxData | Add-Member NoteProperty -Name "Send As Permissions" -Value $SendAsString
	$MailboxData | Add-Member NoteProperty -Name "Send On Behalf Permissions" -Value $SendOnBehalf 
	$MailboxData | Add-Member NoteProperty -Name "Full Access Not Inherited" -Value $FullAccessNotInheritedString

	$MailboxDataOutput += $MailboxData
	}
}

Write-Host "
Exporting mailbox discovery data to CSV file...
"
$MailboxDataOutput | Export-csv -Path $ExportFile -NoTypeInformation
#$MailboxDataOutput | Out-GridView


Write-Host "The report has been created at the following location:
	$ExportFile
" -ForegroundColor DarkCyan