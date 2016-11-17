#########################################################################################
# COMPANY: CDW								                                            #
# NAME: Move-MailboxTo365.ps1                                                           #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  08/28/2014                                                                     #
# EMAIL: Dean.Sesko@S3.CDW.com                                                          #
#                                                                                       #
# COMMENT:  Script to Move Mailboxes to Office 365 				 	                    #
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 08/28/2014 Initial Version.                                                       #
#                                                                                       #
#########################################################################################
param ([Parameter(Mandatory = $true)]
	[string]$UPN
)
$bg = (Get-Host).UI.RawUI
$BG.BackgroundColor = "black"

Try { $test = Get-MsolDomain -ErrorAction stop | where { ($_.name -like "*.mail.onmicrosoft.com") } }
Catch { Invoke-Expression .\Connect365.ps1 }
Finally { }

$TenantRoutingDomain = $TenantDom = Get-MsolDomain | where { $_.name -like "*.mail.onmicrosoft.com" }


Try { $test = get-pssession -Name "ON-Prem-Exchange" -ErrorAction stop }
Catch { Invoke-Expression .\ConnectLocalExchange.ps1 }
Finally { }




$user = get-user $upn -ErrorAction silentlycontinue

if ($user) {
	#Move Mailbox
	
	#if ($user.ExchangeVersion.ExchangeBuild.Major -lt 7) {
		
	#	Write-Host "Source Mailbox is 2003, it will not be AutoSuspended" -ForegroundColor Green -BackgroundColor Black
    #	Write-Host "Moving Mailbox: " $upn -ForegroundColor Green -BackgroundColor Black
	#	new-moverequest -identity $upn -remote -remotehostname $global:HybridServer -targetdeliverydomain $TenantRoutingDomain.name  -remotecredential $global:onPremCred -baditemlimit 10 -LargeItemLimit 10 
	#	Set-MsolUser –UserPrincipalName $upn -UsageLocation "US"
	#}
	#else {
		
		Write-Host "Moving Mailbox: " $upn -ForegroundColor Green -BackgroundColor Black
		new-moverequest -identity $upn -remote -remotehostname $global:HybridServer -targetdeliverydomain $TenantRoutingDomain.name  -remotecredential $global:onPremCred -baditemlimit 10 -LargeItemLimit 10 -SuspendWhenReadyToComplete
		Set-MsolUser –UserPrincipalName $upn -UsageLocation "US"
	#}
	
}



Else {
	
	Write-Host "User Does Not Exist." -ForegroundColor Red -BackgroundColor Black
	
}

