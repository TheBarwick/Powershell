<#
Company: CDW
Name: Office365Discovery.ps1
Author: Matthew Sherwood
Author Email: matt.sherwood@s3.cdw.com
Date Created: 5/1/2016

*** Requirements ***
	Script must be run on system that meets the following requirements
	-List of UPNs in a CSV file
		+ List cannot contain a header. Only a list of UPNs.
	-Microsoft Online Services Sign-In Assistant for IT Professionals RTW
		+ http://go.microsoft.com/fwlink/?LinkID=286152
	-Azure Active Directory Module for Windows PowerShell (64-bit version)
		+ http://go.microsoft.com/fwlink/p/?linkid=236297
	-Server Remote Managment Tools - Active Directory Management Tools
		+ https://technet.microsoft.com/en-us/library/cc816817(v=ws.10).aspx
		
*** Change Log ***
5/4/16
	Added check to see if user is a part of an AD Group existence

#>


#Adjust these variables to fit your needs
#Import file. Specify CSV file that only includes list of UPNs, without a header. 
$ImportFile = ".\UserImport.csv"

#Export File. Specify the directory and file name you would like to save your report in. 
#Default saves report to same directory as script, with date/time info.
$DateTime = "$((Get-Date).ToString('MM-dd-yyyy_hh-mm-ss'))"
$ExportFile = ".\$DateTime Office365Report.csv"
#$ExportFile = Read-Host "Enter the location to save the CSV report (Eg. C:\Scripts\Office365Report.csv)"

# Enter a group name to see if user is a member of this group
$AdGroupMember1 = "Intune_users"
$AdGroupMember2 = "Intune_Provisioning"

#------------------------- Do Not Change Below This Line -------------------------#
#------------------------- Do Not Change Below This Line -------------------------#
#------------------------- Do Not Change Below This Line -------------------------#


#Import modules
Import-Module MSOnline
Import-Module ActiveDirectory

#Create null array
$ReportData = @()
$ReportDataOutput = @()



#Connect to Azure AD
Try {$test = Get-MsolDomain -ErrorAction stop | where { ($_.name -like "*.mail.onmicrosoft.com") } }
Catch{Invoke-Expression .\Connect365.ps1 }
Finally{ }

cls


#Setup Shell
$bg = (Get-Host).UI.RawUI
$BG.BackgroundColor = "Black" 
cls


#Import UPN list
$UPNs = Get-Content $ImportFile
#$UPNs = Get-mailbox -resultsize unlimited | where {$_.RecipientTypeDetails -ne "DiscoveryMailbox"}


Write-Host "
Starting to gather data for each UPN...
"

#------------------------- Functions -------------------------#


Function Test-ADUserExists  {
	Param ($User)
		Trap {Return "error"}
		if (Get-ADUser -Filter {UserPrincipalName -like $User}) {
			$true
		} else {
			$false
		}
}

Function Test-MsolUserExists {
	Param ($User)
		Trap {Return "error"}
		if (Get-MsolUser -UserPrincipalName $User -ErrorAction SilentlyContinue) {
			$true
		} else {
			$false
		}
}

Function Test-ADGroupMember {
	Param ($User,$Group)
		Trap {Return "error"}
		If (Get-ADUser `
			-Filter "memberOf -RecursiveMatch '$((Get-ADGroup $Group).DistinguishedName)'" `
			-SearchBase $((Get-ADUser $User).DistinguishedName) `
			-ErrorAction SilentlyContinue
		) {
			$true
		} else {
			$false
		}
}


#------------------------- Main script -------------------------#
Foreach ($UPN in $UPNs) {
	#Create variables and clear previous variables...
	#General Info
	$ADUser = @()
	$MsolUser = @()
	$MsolMailbox = @()
	$RecipientDisplayType = ""
	$RecipientTypeDetails = ""
	$RemoteRecipientType = ""
	$ADUserExists = @()
	$MsolUserExists = @()
	

	#Create the ReportData object
	$ReportData = New-Object PSObject


	#Gather Data...
	#Verify that UPN exists in AD, and also store AD properties in the ADUser variable
	$ADUserExists = Test-ADUserExists -User $UPN

	
	
	if ($ADUserExists -eq $false) {
		Write-Host "$UPN doesn't exist in AD. Skipping user." -ForegroundColor Red
	} elseif ($ADUserExists -eq $true) {
		Write-Host "$UPN exists in AD" -ForegroundColor Green
		Write-Host "Gathering data about $UPN" -ForegroundColor Green
		
		#Get AD user data
		$ADUser = Get-ADUser -Filter {UserPrincipalName -like $UPN} -Properties * -ErrorAction SilentlyContinue
		
		#Set other variables
		if ($ADUser.msExchRecipientDisplayType -eq "-2147483642") {
			$RecipientDisplayType = "SyncedMailboxUser"
		} else {
			$RecipientDisplayType = $ADUser.msExchRecipientDisplayType
		}

		if ($ADUser.msExchRecipientTypeDetails -eq "2147483648") {
			$RecipientTypeDetails = "RemoteMailbox"
		} else {
			$RecipientTypeDetails = $ADUser.msExchRecipientTypeDetails
		}

		if ($ADUser.msExchRemoteRecipientType -eq "1") {
			$RemoteRecipientType = "ProvisionedMailbox (Cloud MBX)"
		} else {
			$RemoteRecipientType = $ADUser.msExchRemoteRecipientType
		}
	

		#Verify that UPN exists in Office 365, and also store MSOL properties in the MSOL variable
		$MsolUserExists = Test-MsolUserExists -User $UPN
	
		if ($MsolUserExists -eq $false) {
			Write-Host "$UPN doesn't exist in MSOL" -ForegroundColor Red
		} elseif ($MsolUserExists -eq $true) {
			Write-Host "$UPN exists in MSOL" -ForegroundColor Green
			$MsolUser = Get-MsolUser -UserPrincipalName $UPN | select IsLicensed -ErrorAction SilentlyContinue
		} else {
			Write-Host "$UPN MSOL check failed" -ForegroundColor Red
		}
		
		
		$AdGroupMemberCheck1 = Test-ADGroupMember -User $ADUser.DistinguishedName -Group $AdGroupMember1
		$AdGroupMemberCheck2 = Test-ADGroupMember -User $ADUser.DistinguishedName -Group $AdGroupMember2

		#Add first set of Mailbox data properties to the ReportDataOutput object
		$ReportData | Add-Member NoteProperty -Name "Display Name" -Value $ADUser.displayname
		$ReportData | Add-Member NoteProperty -Name "Retention Policy" -Value $MsolMailbox.RetentionPolicy
		$ReportData | Add-Member NoteProperty -Name "Is Licensed" -Value $MsolUser.IsLicensed
		$ReportData | Add-Member NoteProperty -Name "targetAddress" -Value $ADUser.targetAddress
		$ReportData | Add-Member NoteProperty -Name "RecipientTypeDetails" -Value $RecipientTypeDetails
		$ReportData | Add-Member NoteProperty -Name "RecipientDisplayType" -Value $RecipientDisplayType
		$ReportData | Add-Member NoteProperty -Name "homeMDB" -Value $ADUser.homeMDB
		$ReportData | Add-Member NoteProperty -Name "homeMTA" -Value $ADUser.homeMTA
		$ReportData | Add-Member NoteProperty -Name "msExchHomeServerName" -Value $ADUser.msExchHomeServerName
		$ReportData | Add-Member NoteProperty -Name "RemoteRecipientType" -Value $RemoteRecipientType
		$ReportData | Add-Member NoteProperty -Name "User in $AdGroupMember1 AD group" -Value $AdGroupMemberCheck1
		$ReportData | Add-Member NoteProperty -Name "User in $AdGroupMember2 AD group" -Value $AdGroupMemberCheck2
		$ReportData | Add-Member NoteProperty -Name "TestMsolUserExists" -Value $MsolUserExists
		

		<#
		$ReportData | Add-Member NoteProperty -Name "msExchVersion" -Value $ADUser.msExchVersion
		$ReportData | Add-Member NoteProperty -Name "msExchRecipientDisplayType" -Value $ADUser.msExchRecipientDisplayType
		$ReportData | Add-Member NoteProperty -Name "msExchRecipientTypeDetails" -Value $ADUser.msExchRecipientTypeDetails
		$ReportData | Add-Member NoteProperty -Name "msExchRemoteRecipientType" -Value $ADUser.msExchRemoteRecipientType
		#>

		$ReportDataOutput += $ReportData
		
	} else {
		Write-Host "$UPN AD check failed"
	}

}

Write-Host "
Exporting mailbox discovery data to CSV file...
"

$ReportDataOutput | Export-csv -Path $ExportFile -NoTypeInformation
#$ReportDataOutput | Out-GridView


Write-Host "The report has been created at the following location:
	$ExportFile
" -ForegroundColor DarkCyan