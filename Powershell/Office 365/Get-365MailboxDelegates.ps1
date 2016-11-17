#########################################################################################
#                                                                                       #
# COMPANY: CDW                                                                          #
# NAME: Get-365MailboxDelegates.ps1                                                     #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  06/11/2015                                                                     #
# EMAIL: Dean.Sesko@S3.cdw.com                                                          #
#                                                                                       #
# COMMENT:  Script finds users delegate relationships for Full Access, Send as and      #
#           Send on Behalf of Access													#
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 06/11/2015 Initial Version.                                                       #
#                                                                                       #
#########################################################################################

#Variable of script Directory
$ScriptDir = "C:\Scripts\"


Cls

#------------------------ Do not Modify Below this line-------------------------------------------------------------------------#
#Setup UI
$bg = (Get-Host).UI.RawUI
$BG.BackgroundColor = "black"
# Seup Variables
$records = @()
$mailboxData = @()

# Export File for Mailbox and Delegate relationship
$delegateOutputfileName = "Delegates.csv"

#Export File used to import mailbox information
$mailboxExportFileName = "Mailboxinfo.csv"


#DataTable setup
$mailboxinfoTable = New-Object System.Data.DataTable
$DelegateTable = New-Object System.Data.DataTable

$Sdate = get-date
$index = 0
$Global:output = @()
$global:FAdelegateCount = 0
$global:SADelegateCount = 0
$global:SOBDelegateCount = 0
[Boolean]$global:DBFindSetup = $false


#Setup Logs
$DelegateOutput = "$ScriptDir$delegateOutputfileName"
$MailboxOutPut = "$ScriptDir$mailboxExportFileName"

Function FindDelegatesStartUp {
	$Sdate = get-date
	cls
	Write-Host ""
	Write-Host ""
	Write-Host ""
	Write-Host ""
	Write-Host ""
	Write-Host ""
	Write-Host "Connecting to Office 365 and getting a list of Mailboxes.  This Process can a long time in large environments." -ForegroundColor Green -BackgroundColor Black
	Write-Host ""
	Write-Host "Get-Mailbox Start Time: $sdate"
	#End Function
}
Function FindDelegatesCloseOut {
	$EDate = get-date
	$mbxTime = $MyMBXPERSEC.trim()
	Write-Host ""
	Write-Host ""
	Write-Host "Finsihed Gathering Delegate Information" -ForegroundColor Green -BackgroundColor Black
	Write-Host ""
	Write-Host "-------------------------------------------------------------"
	Write-Host "Process Started at:  $sdate  "
	Write-Host "Process ended at:  $EDate  "
	Write-Host ""
	Write-Host "Found $global:FAdelegateCount Delegates with Full Access Permissions  "
	Write-Host "Found $SAdelegateCount Delegates with Send As Permissions"
	Write-Host "Found $global:SOBDelegateCount Delegates with Send on Be Half Of Permissions  "
	Write-Host ""
	Write-Host "Mailboxes Per Second $mbxTime"
	Write-Host "-------------------------------------------------------------"
	Write-Host ""
	$DelegateTable | Export-csv $DelegateOutput -NoTypeInformation
	$mailboxinfoTable | Export-csv $MailboxOutPut -NoTypeInformation
	Write-Progress -Activity "   $ProcessData   " -status "Ready" -Completed
	
	
	#End Function
}
Function SetupFindDelegatesDB {
	$mailboxinfoTable.Columns.Add("SamAccountName") | Out-Null
	$mailboxinfoTable.Columns.Add("PrimarySMTPAddress") | Out-Null
	$mailboxinfoTable.Columns.Add("Alias") | Out-Null
	$pk = $mailboxinfoTable.Columns["SamAccountName"]
	$pk1 = $mailboxinfoTable.Columns["Alias"]
	$mailboxinfoTable.PrimaryKey = $pk, $pk1
	
	$DelegateTable.Columns.Add("MailBox") | Out-Null
	$DelegateTable.Columns.Add("Delegate") | Out-Null
	$DelegateTable.Columns.Add("Permission") | Out-Null
	
	$global:DBFindSetup = $true
	#End Function
}
Function FillDelegatetable($mailbox, $delegate, $perm) {
	
	$row = $DelegateTable.NewRow()
	$row["MailBox"] = $mailbox
	$row["Delegate"] = $delegate
	$row["Permission"] = $perm
	$DelegateTable.Rows.Add($row)
	
	#End Function
}
Function GetSmtpAddress($mbx) {
	if ($mbx.primarySMTPAddress -eq $null) {
		$mbx.EmailAddresses | ForEach {
			If ($_.Prefix -match "SMTP") {
				$SmtpAddress = $_.SmtpAddress
			}
		}
	}
	
	Else {
		$SmtpAddress = $mbx.primarysmtpAddress.tostring()
	}
	return $SmtpAddress
	#End Function
}
Function UpdateProgressBar($Users, $ProgressIndex, $Progress, $ProcessData) {
	$progress = ($ProgressIndex/$Users) * 100
	$FP = "{0:N2}" -f $progress
	$EDate = get-date
	$seconds = (New-TimeSpan -Start $MbxProcStartTime -End $EDate).totalseconds
	$MBXPerSec = $seconds / $ProgressIndex
	$MBXPS = "{0:N2}" -f $MBXPerSec
	Write-Progress -Activity "   $ProcessData   " -status "   $FP% Complete: Processing $MBXPS Seconds/Mailbox: " -PercentComplete $FP
	Return $MBXPS
	#End Function
}
Function GetFullAccessDelegate($mbx) {
	$results = Get-MailboxPermission $MBX.Name -erroraction Silentlycontinue -WarningAction silentlyContinue | Where { ($_.AccessRights -eq “FullAccess”) -and -not ($_.isInherited -like "True") } | Select User
	if ($results) {
		$global:FAdelegateCount++
		foreach ($result in $results) {
			$Delegate = $result.user
			FillDelegatetable $MBX.PrimarySMTPAddress $Delegate "FullAccess"
			
		}
	}
	#End FUnction
}
Function GetSendONDelegate($mbx) {
	$delegateINfo = @()
	$global:SOBDelegateCount++
	foreach ($delegate in $mbx.GrantSendOnBehalfTo) {
		$RowDelegate = $null
		$MyDelegate = $null
		[string]$filter = " Alias = '" + $delegate + "'"
		Try {
			$RowDelegate = $mailboxinfoTable.Select($filter)
			$MyDelegate = $RowDelegate[0]["PrimarySMTPAddress"].ToString()
		}
		Catch { }
		Finally { }
		FillDelegatetable $MBXAddress $MyDelegate "SendonBeHalfOf"
	}
	#End Function
}
Function GetSendAsDelegate($mbx) {
	
	$results = Get-RecipientPermission $MBX.Name -erroraction Silentlycontinue -WarningAction silentlyContinue |Where { ($_.AccessRights -eq "SendAs") -and  ($_.Trustee -ne "NT AUTHORITY\SELF") } | Select Trustee
	if ($results) {
		$global:SADelegateCount++
		foreach ($result in $results) {
			$Delegate = $result.Trustee
			FillDelegatetable $MBX.PrimarySMTPAddress $Delegate "SendAs"
			
		}
	}
	
}
Function GetAllMailboxes {
	$Mailboxes = get-mailbox -resultsize unlimited -WarningAction silentlyContinue | select primarySMTPAddress, Database, Name, ServerName, GrantSendOnBehalfTo, SamAccountName, OrganizationalUnit,Alias
	Return $Mailboxes
	#End Function
}
Function ProcessDelegates {
	$Progress = 0
	$ProgressIndex = 0
	$ProcessData = " Importing  Mail Data"
	foreach ($MBX in $MyMailboxes) {
		$MBXAddress = GetSmtpAddress($mbx)
		$FullAccessDelegateSMTP = GetFullAccessDelegate($mbx)
		if ($mbx.GrantSendOnBehalfTo) {
			$SendonBehalfoDelegateSMTP = GetSendONDelegate($mbx)
		}
		
		GetSendAsDelegate($mbx)
		$ProgressIndex++
		$ProgressUpdate = UpdateProgressBar $MyMailboxes.count $Progressindex $Progress $ProcessData
		
		#End Function
	}
	
	Return $ProgressUpdate
}

Function FillMailboxinfoDataTable {
	
	foreach ($MBX in $MyMailboxes) {		
		$MBXAddress = GetSmtpAddress($mbx)
		$row = $mailboxinfoTable.NewRow()
		$row["SamAccountName"] = $MBX.SamAccountName
		$row["PrimarySMTPAddress"] = $MBXAddress
		$row["Alias"] = [string]$path = $MBX.alias
		$mailboxinfoTable.Rows.Add($row)
	}
	
	#End Function
}
Write-Host "Connected to Office 365......" -ForegroundColor Green -BackgroundColor Black
FindDelegatesStartUp
$MyMailboxes = GetAllMailboxes
$MbxProcStartTime = Get-Date


if ($global:DBFindSetup -eq $false) {
	SetupFindDelegatesDB
}
Else {
	
	$DelegateTable.Clear()
}
FillMailboxinfoDataTable
[string]$MyMBXPERSEC = ProcessDelegates
FindDelegatesCloseOut



#End Main


