#########################################################################################
# COMPANY: CDW								                                            #
# NAME: Disable-O365Mailbox.ps1                                                         #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  03/25/2014                                                                     #
# EMAIL: Dean.SEsko@CDW.com                                                             #
#                                                                                       #
# COMMENT:  Script to Remove an Office 365 mailbox from and Existing user               #
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 03/25/2014 Initial Version.                                                       #
# 1.1 10/20/2014 Renamed Script and worked on script format and connection methods      #
#                                                                                       #
#########################################################################################
param ( [Parameter(Mandatory=$true)] 
  [string]$UPN
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
        Try{
	    Write-Host "Disabling Office 365 Mailbox: " $upn  -ForegroundColor Green -BackgroundColor Black
		Disable-onPremRemoteMailbox -Identity $user.UserPrincipalName -erroraction stop -confirm:$false
        Invoke-Expression .\StartSync.ps1 
        Write-host ""
        Write-host ""
        Write-Host $upn"'s Office 365 mailbox has been disabled"  -ForegroundColor Green -BackgroundColor Black
        Write-host "Please login to the Office 365 portal and remove the License from the user account if it is no longer needed" -ForegroundColor Green -BackgroundColor Black

         }

         Catch { Write-Host "Error Disabling Office 365 Mailbox for User: " $upn  -ForegroundColor Red -BackgroundColor Black
                 Write-Host "The Users mailbox may not exist or there is an error connection to local Directory services"  -ForegroundColor Red -BackgroundColor Black }
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



