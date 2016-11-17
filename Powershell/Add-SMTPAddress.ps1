#########################################################################################
# COMPANY: CDW								                                            #
# NAME: Add-SMTPAddress.ps1                                                             #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  03/25/2014                                                                     #
# EMAIL: Dean.SEsko@CDW.com                                                             #
#                                                                                       #
# COMMENT:  Script to add an additional email address to an Office 365 mailbox          #
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 10/20/2014 Initial Version.                                                       #
#                                                                                       #
#########################################################################################
param ( [Parameter(Mandatory=$true)] 
 [string]$UPN,
 [Parameter(Mandatory=$true)] 
 [string]$EmailAddress
)
#	Setup Shell
$bg = (Get-Host).UI.RawUI
$BG.BackgroundColor = "black"
cls

# Check for Local Exchnage Connection.  If fails make a connection
Try {$test=get-pssession -Name "ON-Prem-Exchange" -ErrorAction stop}
Catch{Invoke-Expression .\ConnectLocalExchange.ps1 }
Finally{}




if (($UPN)){
Write-Host ""
Write-Host ""
    	
    $user = get-onpremuser $upn -ErrorAction silentlycontinue
    if ($user){

    Try {set-onPremRemoteMailbox $UPN -EmailAddresses @{ add = $EmailAddress } -ErrorAction Stop -WarningAction silentlycontinue
    Write-host "Email Address for user:$upn has been Added" -ForegroundColor Green -BackgroundColor Black}
    Catch {Write-Host "Error Updating Email Addresses"}
    Finally{}
    }
    Else{
		 CLS
         Write-Host""
         Write-Host""
         Write-Host "User Does Not Exist." -ForegroundColor Red -BackgroundColor Black
         Write-Host""
         Write-Host""
	}

}




