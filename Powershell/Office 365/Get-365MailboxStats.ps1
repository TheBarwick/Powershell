#########################################################################################
# COMPANY: CDW                                                                          #
# NAME: Get-365MailboxStats.ps1                                                         #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  06/10/2015                                                                     #
# EMAIL: Dean.SEsko@S3.cdw.com                                                          #
#                                                                                       #
# COMMENT:  Script to Gather Office 365 Mailbox Statistics                              #
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 06/10/2015 Initial Version.                                                       #
#                                                                                       #
#                                                                                       #
#########################################################################################
$OutputPath = "c:\scripts\MailData.csv"

#Create DataTable for Formated Data
$DataTable = New-Object System.Data.DataTable
$DataTable.Columns.Add("DisplayName") | Out-Null
$DataTable.Columns.Add("PrimarySMTPAddress") | Out-Null
$DataTable.Columns.Add("itemcount") | Out-Null
$DataTable.Columns.Add("ItemSizeinMB") | Out-Null

#Setup UI
$bg = (Get-Host).UI.RawUI
$BG.BackgroundColor = "black"
Cls
Write-Host ""
Write-Host "Collecting Mailbox Data" 

$Mailboxes = get-mailbox -ResultSize unlimited
Write-Host "Found"  $Mailboxes.count  "Mailboxes" -ForegroundColor 'Green'

$index =0
foreach ($mbx in $Mailboxes) {
	#Update Progress 
	$index ++
	$progress = ($index / $Mailboxes.count) * 100
	$progress = [math]::Round($progress, 2)
	Write-Progress -Activity "   Gathering Mailbox Data   " -Status " $progress% " -PercentComplete $progress
	
	#Get Mailbox progress 
	$stat = Get-MailboxStatistics $mbx.PrimarySMTPAddress
	#Populate Data Table
	$row = $DataTable.NewRow()
	$row["DisplayName"] = $mbx.DisplayName
	$row["PrimarySMTPAddress"] = $mbx.PrimarySMTPAddress
	$row["itemcount"] = $stat.itemcount
	$ConvertedtoMB = $stat.TotalItemSize.toString().split("(")
	$ConvertedtoMB = $ConvertedtoMB[1].Split(" bytes)")
	$ConvertedtoMB = $ConvertedtoMB[0].trim()
	$ConvertedtoMB = $ConvertedtoMB -replace ',', ''
	$ConvertedtoMB = (($ConvertedtoMB)/1024)/1024
	$ConvertedtoMB = [math]::Round($ConvertedtoMB, 2)
	$row["ItemSizeinMB"] = $ConvertedtoMB
	$DataTable.Rows.Add($row)
	
}
#Export Data Table
$DataTable | export-csv -path $OutputPath -NoTypeInformation

