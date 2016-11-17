#########################################################################################
# COMPANY: CDW                                                                          #
# NAME: LicenseReport.ps1                                                               #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  09/22/2015                                                                     #
#                                                                                       #
# EMAIL: Dean.Sesko@s3.cdw.com                                                          #
#                                                                                       #
# COMMENT:  Script easily report on users licenses                                      #
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 09/22/2015 Initial Version.                                                       #
#                                                                                       #
#########################################################################################
param ([Parameter(Mandatory = $false)]	[bool]$HTMLReport =$false,
	   [Parameter(Mandatory = $false)]	[bool]$CSVReport = $false
)
$bg = (Get-Host).UI.RawUI
$BG.BackgroundColor = "black"

Try { $test = Get-MsolDomain -ErrorAction stop | where { ($_.name -like "*.mail.onmicrosoft.com") } }
Catch { Invoke-Expression .\Connect365.ps1 }
Finally { }

#$MSOLUsers = Get-MsolUser -All

$LicTable = New-Object System.Data.DataTable
$LicTable.Columns.Add("UserPrincipalName") | Out-Null
$LicTable.Columns.Add("IsLicensed") | Out-Null
$LicTable.Columns.Add("License") | Out-Null
$LicTable.Columns.Add("INTUNE_O365") | Out-Null
$LicTable.Columns.Add("YAMMER_ENTERPRISE") | Out-Null
$LicTable.Columns.Add("MCOSTANDARD") | Out-Null
$LicTable.Columns.Add("RMS_S_ENTERPRISE") | Out-Null
$LicTable.Columns.Add("OFFICESUBSCRIPTION") | Out-Null
$LicTable.Columns.Add("SHAREPOINTENTERPRISE") | Out-Null
$LicTable.Columns.Add("EXCHANGE_S_ENTERPRISE") | Out-Null
$LicTable.Columns.Add("SHAREPOINTSTANDARD") | Out-Null
$LicTable.Columns.Add("EXCHANGE_S_STANDARD") | Out-Null
$LicTable.Columns.Add("SHAREPOINTWAC") | Out-Null




function GenerateHTMLPage ($DT) {
	$Output = "
<html>
<body>
<font size=""1"" face=""Arial,sans-serif"">
<h1 align=""center"">Office 365 License Report</h1>
<h3 align=""center"">Generated $((Get-Date).ToString())</h3>
</font><br>"
	
	
	
	$Output += "<table border=""2"" cellpadding=""0"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif"">"
	$Output += "<tr><th><font color=""#000000""> <h3> User Principal Name </h3></font></th>
	<th><font color=""#000000""><h3>Has 365 License</h3></font></th>
	<th><font color=""#000000""><h3>Assigned License</h3></font></th>
	<th><font color=""#000000""><h3>ITUNE License</font></h3></th>
	<th><font color=""#000000""><h3>Yammer License</font></h3></th>
	<th><font color=""#000000""><h3>Skype For Business License</font></h3></th>
	<th><font color=""#000000""><h3>RMS License</font></h3></th>
	<th><font color=""#000000""><h3>Office Pro Plus</font></h3></th>
	<th><font color=""#000000""><h3>SharePoint Enterprise License</font></h3></th>
	<th><font color=""#000000""><h3>Exchange Enterprise License</font></h3></th>
	<th><font color=""#000000""><h3>Sharepoint Standard License</font></h3></th>
	<th><font color=""#000000""><h3>Exchange Standard License</font></h3></th>
	<th><font color=""#000000""><h3>SharePoint Web Application License</h3></font></th></tr>"
	$Output += "</tr><tr><tr>"
	
	foreach ($User in $DT) {
		$Output += "<tr "
		if ($AlternateRow) {
			$Output += " style=""background-color:#dddddd"""
			$AlternateRow = $false
		}
		else {
			$AlternateRow = $true
		}
		
		$Output += "><td>$($User.Userprincipalname)"
		$Output += "<td>$($User.IsLicensed)"
		$Output += "<td>$($User.License)"
		if ($User.INTUNE_O365 -ne "Disabled") { $Output += "<td>$($User.INTUNE_O365)" }
		Else { $Output += "<td style=""background-color:#FF0000""> $($User.INTUNE_O365)" }
		if ($User.YAMMER_ENTERPRISE -ne "Disabled") { $Output += "<td>$($User.YAMMER_ENTERPRISE)" }
		Else { $Output += "<td style=""background-color:#FF0000""> $($User.YAMMER_ENTERPRISE)" }
		if ($User.MCOSTANDARD -ne "Disabled") { $Output += "<td>$($User.MCOSTANDARD)" }
		Else { $Output += "<td style=""background-color:#FF0000""> $($User.MCOSTANDARD)" }
		if ($User.RMS_S_ENTERPRISE -ne "Disabled") { $Output += "<td>$($User.RMS_S_ENTERPRISE)" }
		Else { $Output += "<td style=""background-color:#FF0000""> $($User.RMS_S_ENTERPRISE)" }
		if ($User.OFFICESUBSCRIPTION -ne "Disabled") { $Output += "<td>$($User.OFFICESUBSCRIPTION)" }
		Else { $Output += "<td style=""background-color:#FF0000""> $($User.OFFICESUBSCRIPTION)" }
		if ($User.SHAREPOINTENTERPRISE -ne "Disabled") { $Output += "<td>$($User.SHAREPOINTENTERPRISE)" }
		Else { $Output += "<td style=""background-color:#FF0000""> $($User.SHAREPOINTENTERPRISE)" }
		if ($User.EXCHANGE_S_ENTERPRISE -ne "Disabled") { $Output += "<td>$($User.EXCHANGE_S_ENTERPRISE)" }
		Else { $Output += "<td style=""background-color:#FF0000""> $($User.EXCHANGE_S_ENTERPRISE)" }
		if ($User.SHAREPOINTSTANDARD -ne "Disabled") { $Output += "<td>$($User.SHAREPOINTSTANDARD)" }
		Else { $Output += "<td style=""background-color:#FF0000""> $($User.SHAREPOINTSTANDARD)" }
		if ($User.EXCHANGE_S_STANDARD -ne "Disabled") { $Output += "<td>$($User.EXCHANGE_S_STANDARD)" }
		Else { $Output += "<td style=""background-color:#FF0000""> $($User.EXCHANGE_S_STANDARD)" }
		if ($User.SHAREPOINTWAC -ne "Disabled") { $Output += "<td>$($User.SHAREPOINTWAC)" }
		Else { $Output += "<td style=""background-color:#FF0000""> $($User.SHAREPOINTWAC)" }
		$Output += "</tr>"
	}
	
	$Output += "</table>"
	$Output += "</body></html>";
	
	return $Output
}

foreach ($MSOLUser in $MSOLUsers) {
	$row = $LicTable.NewRow()
	$row["UserPrincipalName"] = $MSOLUser.UserPrincipalname
	
	
	if ($MSOLUser.IsLicensed) {
		$row["IsLicensed"] = "True"
		foreach ($License in $MSOLUser.Licenses) {
			
			
			$row["License"] = $License.AccountSku.SkuPartNumber
			
			if ($License.AccountSku.SkuPartNumber -eq "ENTERPRISEPACK") {
				
				foreach ($plan in $License.ServiceStatus) {
					
					switch ($plan.Serviceplan.Servicename) {
						
						"INTUNE_O365"{ $row["INTUNE_O365"] = $plan.ProvisioningStatus }
						"YAMMER_ENTERPRISE"{ $row["YAMMER_ENTERPRISE"] = $plan.ProvisioningStatus }
						"RMS_S_ENTERPRISE"{ $row["RMS_S_ENTERPRISE"] = $plan.ProvisioningStatus }
						"OFFICESUBSCRIPTION"{ $row["OFFICESUBSCRIPTION"] = $plan.ProvisioningStatus }
						"MCOSTANDARD"{ $row["MCOSTANDARD"] = $plan.ProvisioningStatus }
						"SHAREPOINTWAC"{ $row["SHAREPOINTWAC"] = $plan.ProvisioningStatus }
						"SHAREPOINTENTERPRISE"{ $row["SHAREPOINTENTERPRISE"] = $plan.ProvisioningStatus }
						"EXCHANGE_S_ENTERPRISE"{ $row["EXCHANGE_S_ENTERPRISE"] = $plan.ProvisioningStatus }
					}
					
				}
				$row["SHAREPOINTSTANDARD"] = "NA"
				$row["EXCHANGE_S_STANDARD"] = "NA"
			}
			
			elseif ($License.AccountSku.SkuPartNumber -eq "STANDARDPACK") {
				foreach ($plan in $License.ServiceStatus) {
					switch ($plan.Serviceplan.Servicename) {
						
						"INTUNE_O365"{ $row["INTUNE_O365"] = $plan.ProvisioningStatus }
						"YAMMER_ENTERPRISE"{ $row["YAMMER_ENTERPRISE"] = $plan.ProvisioningStatus }
						"SHAREPOINTSTANDARD"{ $row["SHAREPOINTSTANDARD"] = $plan.ProvisioningStatus }
						"EXCHANGE_S_STANDARD"{ $row["EXCHANGE_S_STANDARD"] = $plan.ProvisioningStatus }
						"MCOSTANDARD"{ $row["MCOSTANDARD"] = $plan.ProvisioningStatus }
						
						
					}
				}
				$row["SHAREPOINTWAC"] = "NA"
				$row["SHAREPOINTENTERPRISE"] = "NA"
				$row["EXCHANGE_S_ENTERPRISE"] = "NA"
				$row["RMS_S_ENTERPRISE"] = "NA"
				$row["OFFICESUBSCRIPTION"] = "NA"
			}
			
		}
	}
	
	Else {
		$row["IsLicensed"] = "False"
		$row["License"] = "NA"
	}
	
	
	$LicTable.Rows.Add($row)
}


$SortedView = New-Object System.Data.DataView($LicTable)
$SortedView.Sort = "UserPrincipalName"
$SortedView | FT

if ($HTMLReport) {
	$HtmlOutPut =GenerateHTMLPage $SortedView
	$HtmlOutPut | Out-File LicenseReport.html
}

if ($CSVReport) {
	$SortedView | Export-Csv C:\Scripts\LicenseReport.csv -NoTypeInformation
}