#########################################################################################
# COMPANY: CDW                                                                          #
# NAME: Add-BulkDomain365.ps1                                                           #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  09/25/2015                                                                     #
#                                                                                       #
# EMAIL: Dean.Sesko@s3.cdw.com                                                          #
#                                                                                       #
# COMMENT:  Add Domains to Office 365 in Bulk                                           #
#           Sript requires a csv file with 1 record per row and a header of DomainFQDN  #
#           The Default Location is C:\scripts                                          #
#           The script requires a remote powershell connection to Office365 with the    #
#           MSONLINE Modules loaded                                                     #
#                                                                                       #
# SAMPLE CSV                                                                            #
#           DomainFQDN                                                                  #
#           contoso.com                                                                 #
#           Tailspintoys.com                                                            #
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 09/25/2015 Initial Version.                                                       #
#                                                                                       #
#########################################################################################
$Domains = Import-Csv C:\Scripts\Domains.Csv

$results = New-Object System.Data.DataTable
$results.Columns.Add("DomainFqdn") | Out-Null
$results.Columns.Add("DNSTextRecord") | Out-Null


foreach ($NewDomain in $Domains) {
	
	New-MsolDomain -Name $NewDomain.DomainFqdn
	$DNSTxt = Get-MsolDomainVerificationDns -DomainName $NewDomain.DomainFqdn
	
	$row = $results.NewRow()
	$row["DomainFqdn"] = $NewDomain.DomainFqdn
	$row["DNSTextRecord"] = $DNSTxt.label.split(".")[0]
	$results.Rows.Add($row)
}
$results | Export-Csv C:\Scripts\DomainProof.csv -NoTypeInformation
