#########################################################################################
# COMPANY: CDW								                                            #
# NAME: Move-MultipleMailboxTo365.ps1                                                   #
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
$bg = (Get-Host).UI.RawUI
$BG.BackgroundColor = "black"

$Users = Get-Content C:\Scripts\Users.csv

#Test Office 365 Connection 
Try { $test = Get-MsolDomain -ErrorAction stop | where { ($_.name -like "*.mail.onmicrosoft.com") } }
Catch { Invoke-Expression .\Connect365.ps1 }
Finally { }

$TenantRoutingDomain = $TenantDom = Get-MsolDomain | where { $_.name -like "*.mail.onmicrosoft.com" }

#Test Office On Prem Connection 
Try { $test = get-pssession -Name "ON-Prem-Exchange" -ErrorAction stop }
Catch { Invoke-Expression .\ConnectLocalExchange.ps1 }
Finally { }



foreach ($upn in $Users) {
	
	$user = get-user $upn -ErrorAction silentlycontinue
	if ($user) {
		#Move Mailbox
		
		Write-Host "Moving Mailbox: " $upn -ForegroundColor Green -BackgroundColor Black
		new-moverequest -identity $upn -remote -remotehostname $global:HybridServer -targetdeliverydomain $TenantRoutingDomain.name  -remotecredential $global:onPremCred -baditemlimit 10 -LargeItemLimit 10 -SuspendWhenReadyToComplete
		Set-MsolUser –UserPrincipalName $upn -UsageLocation "US"
		
	}
	
	
	
	Else {
		
		Write-Host "User Does Not Exist." -ForegroundColor Red -BackgroundColor Black
		
	}
	
}