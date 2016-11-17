#########################################################################################
# COMPANY: CDW                                                                          #
# NAME: MigrationLicenses.ps1                                                           #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  09/22/2015                                                                     #
#                                                                                       #
# EMAIL: Dean.Sesko@s3.cdw.com                                                          #
#                                                                                       #
# COMMENT:  Script easily report on users licenses and migration progress               #
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 09/22/2015 Initial Version.                                                       #
#                                                                                       #
#########################################################################################
$MSOLUsers = Get-MsolUser -All
$mailboxinfoTable = New-Object System.Data.DataTable
$mailboxinfoTable.Columns.Add("UserPrincipalName") | Out-Null
$mailboxinfoTable.Columns.Add("IsLicensed") | Out-Null
$mailboxinfoTable.Columns.Add("License") | Out-Null
$mailboxinfoTable.Columns.Add("MailboxMoveStatus") | Out-Null
$mailboxinfoTable.Columns.Add("MailboxMovePercentageComplete") | Out-Null
foreach ($MSOLUser in $MSOLUsers) {
	$row = $mailboxinfoTable.NewRow()
	$row["UserPrincipalName"] = $MSOLUser.UserPrincipalname
	
	
	if ($MSOLUser.IsLicensed) {
		foreach ($License in $MSOLUser.Licenses) {
			
			$row["IsLicensed"] = "True"
			$row["License"] = $License.AccountSkuId
			
		}
	}
	
	Else {
		$row["IsLicensed"] = "False"
		$row["License"] = "NA"
	}
	
	Try {
		$moveStats = Get-MoveRequestStatistics $MSOLUser.UserPrincipalname -ErrorAction silentlycontinue
	}
	
	Catch {
		$moveStats = $null
	}
	
	Finally {
		
		if ($moveStats) {
			
			$row["MailboxMoveStatus"] = $moveStats.StatusDetail
			$row["MailboxMovePercentageComplete"] = $moveStats.PercentComplete
			
		}
		Else {
			$row["MailboxMoveStatus"] = "NA"
			$row["MailboxMovePercentageComplete"] = "NA"
			
		}
		
		$mailboxinfoTable.Rows.Add($row)
	}
}

$mailboxinfoTable |FT
Export-Csv C:\Scripts\mailboxMoveinfo.csv -NoTypeInformation

