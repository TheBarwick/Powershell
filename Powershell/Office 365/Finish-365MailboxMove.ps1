#########################################################################################
# COMPANY: CDW                                                                          #
# NAME: Finish-365MailboxMove.ps1                                                       #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  08/28/2014                                                                     #
# EMAIL: Dean.Sesko@S3.CDW.com                                                          #
#                                                                                       #
# COMMENT:  Finishes a suspended Mailbox Move                                           #
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 11/28/2014 Initial Version.                                                       #
#                                                                                       #
#########################################################################################

param ([Parameter(Mandatory = $true)]
	[string]$UPN
)

$bg = (Get-Host).UI.RawUI
$BG.BackgroundColor = "black"

Try { $test = get-pssession -Name "ON-Prem-Exchange" -ErrorAction stop }
Catch { Invoke-Expression .\ConnectLocalExchange.ps1 }
Finally { }

Try { $test = Get-MsolDomain -ErrorAction stop | where { ($_.name -like "*.mail.onmicrosoft.com") } }
Catch { Invoke-Expression .\Connect365.ps1 }
Finally { }


$user = Get-MsolUser -UserPrincipalname $upn -ErrorAction silentlycontinue

if ($user)
{
	#Move Mailbox
	
	Write-Host "Completeing Mailbox Move for Mailbox: " $upn -ForegroundColor Green -BackgroundColor Black
	Get-MoveRequest $upn |Resume-MoveRequest
	Set-MsolUser –UserPrincipalName $upn -UsageLocation "US"
	
}



Else
{
	
	Write-Host "User Does Not Exist." -ForegroundColor Red -BackgroundColor Black
	
}

