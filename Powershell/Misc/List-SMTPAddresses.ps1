#########################################################################################
# COMPANY: CDW								                                            #
# NAME: List-SMTPAddresses.ps1                                                          #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  03/25/2014                                                                     #
# EMAIL: Dean.SEsko@CDW.com                                                             #
#                                                                                       #
# COMMENT:  Script to List email addresses on an Office 365 mailbox                     #
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 10/20/2014 Initial Version.                                                       #
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
		$mbx = $null
		
		Try {
			if ($mbx -eq $null) {
				$mbx = get-onPremMailbox $UPN -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Select-object UserPrincipalName, EmailAddresses | out-null
			}
			if ($mbx -eq $null) {
				$mbx = get-onPremRemoteMailbox $UPN -ErrorAction Stop -WarningAction SilentlyContinue | Select-object UserPrincipalName, EmailAddresses | Out-null
			}
			
			
			cls
            Write-Host "User:" $mbx.UserPrincipalName -ForegroundColor Green
            Write-host ""
            Write-Host "The User has" $mbx.EmailAddresses.Count "Email Aliases" -ForegroundColor Green
            Write-host " "
            Write-host "Email Aliases"
            Write-host "----------------------------------------------------------------------------------"
            foreach($Address in $mbx.EmailAddresses){
            if ($Address.Contains("SMTP")){Write-host $address "(Primary SMTP Address)" -ForegroundColor Yellow} 
            Else{Write-host $address}
            }
            Write-Host ""
     }

     
      
     Catch { 
             Write-Host ""
             Write-Host ""
             Write-Host "User Does not have a Mailbox or the Mailbox was not created properly" -ForegroundColor Red -BackgroundColor Black
             Write-Host "Please run New-365Mailbox -UPN $upn " -ForegroundColor Red -BackgroundColor Black
             Write-Host "This Command will correct the issue or create the mailbox for the specified user"    -ForegroundColor Red -BackgroundColor Black
             Write-Host ""
            }

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




