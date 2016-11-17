#########################################################################################
# COMPANY: CDW                                                                          #
# NAME: Get-365EmailAddresses.ps1                                                       #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  06/10/2015                                                                     #
# EMAIL: Dean.SEsko@S3.cdw.com                                                          #
#                                                                                       #
# COMMENT:  Script to Gather Office 365 Mailbox Email Addresses                         #
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 06/10/2015 Initial Version.                                                       #
#                                                                                       #
#                                                                                       #
#########################################################################################
$OutputPath = "c:\scripts\SMTP.csv"


#Setup UI
$bg = (Get-Host).UI.RawUI
$BG.BackgroundColor = "black"



#Create DataTable for Formated Data
$DataTable = New-Object System.Data.DataTable
$DataTable.Columns.Add("DisplayName") | Out-Null
$DataTable.Columns.Add("PrimarySMTPAddress") | Out-Null
$DataTable.Columns.Add("EmailAddresses") | Out-Null

Cls
Write-Host ""
Write-Host "Collecting Mailbox Data" 

$Mailboxes = Get-DistributionGroup -ResultSize unlimited
Write-Host "Found"  $Mailboxes.count  "Mailboxes" -ForegroundColor 'Green'

$index =0
foreach ($mbx in $Mailboxes) {

	[String]$ua = $null
	Foreach ($add in $mbx.EmailAddresses) {
		if ($add -like "smtp*") {
			$UA += "$add,"
		}
	}
	
	#Update Progress 
	$index ++
	$progress = ($index / $Mailboxes.count) * 100
	$progress = [math]::Round($progress, 2)
	Write-Progress -Activity "   Gathering Mailbox Data   " -Status " $progress% " -PercentComplete $progress
	
	#Get Mailbox progress 
	#Populate Data Table
	$row = $DataTable.NewRow()
	$row["DisplayName"] = $mbx.DisplayName
	$row["PrimarySMTPAddress"] = $mbx.PrimarySMTPAddress
	$row["EmailAddresses"] = $ua
	$DataTable.Rows.Add($row)
	
}
#Export Data Table
$DataTable | export-csv -path $OutputPath -NoTypeInformation

