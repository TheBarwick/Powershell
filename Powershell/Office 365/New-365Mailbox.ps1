#########################################################################################
# COMPANY: CDW                                                                          #
# NAME: New-365Mailbox.ps1                                                              #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  03/25/2014                                                                     #
# EMAIL: Dean.SEsko@CDW.com                                                             #
#                                                                                       #
# COMMENT:  Script to Create a new Office 365 mailbox for an existing user              #
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 03/25/2014 Initial Version.                                                       #
# 1.1 10/17/2014 Foramt Cleanup and connection detection change.                        #
#                                                                                       #
#                                                                                       #
#########################################################################################
param ([Parameter(Mandatory = $true)]
	[string]$UPN
)
#	Setup Shell
$bg = (Get-Host).UI.RawUI
$BG.BackgroundColor = "black"
$RetryCount = 10
$index = 0
cls


# Try To Connect to Online Tenant.  If fails make a connction

Try { $test = Get-MsolDomain -ErrorAction stop | where { ($_.name -like "*.mail.onmicrosoft.com") } }
Catch { Invoke-Expression .\Connect365.ps1 }
Finally { }

# Check for Local Exchnage Connection.  If fails make a connection
Try { $test = get-pssession -Name "ON-Prem-Exchange" -ErrorAction stop }
Catch { Invoke-Expression .\ConnectLocalExchange.ps1 }
Finally { }


#set Script Variables
$TenantRoutingDomain = Get-MsolDomain | where { ($_.name -like "*.mail.onmicrosoft.com") }
$RoutingAddress = $null

#Graphical Sleep Function...  Thanks to Pat Richards @ http://www.ehloworld.com/878
function GSleep ($Time) {
	for ($i = 1; $i -lt $Time; $i++) {
		[int]$TimeLeft = $Time - $i
		Write-Progress -Activity "Waiting $Time seconds..." -PercentComplete (100/$Time * $i) -CurrentOperation "$TimeLeft seconds left ($i elapsed)" -Status "Please wait"
		Start-Sleep -s 1
	}
	Write-Progress -Completed $true -Status "Please wait"
} # end function New-Sleep


#Check to see if UPN is Valid in Office 365
Function CheckOnlineUser {
	
	$user = Get-MsolUser -UserPrincipalName $upn -ErrorAction silentlycontinue
	if ($user -ne $null) {
		Set-MsolUser -UserPrincipalName $upn -UsageLocation "US"
		return $true
	}
	Else {
		Invoke-Expression .\StartSync.ps1
		Write-Host "Checking for Office 365 User:" $UPN -ForegroundColor Green -BackgroundColor Black
		GSleep 30
		Return $false
	}
}


if (($UPN)) {
	
	Write-Host " "
	Write-Host "Attempting to Connect to Local User: " $upn -ForegroundColor Green -BackgroundColor Black
	Write-Host " "
	$user = get-OnPremUser $upn -ErrorAction silentlycontinue
	if ($user) {
		if (!(Get-OnPremRemoteMailbox $upn -erroraction silentlyContinue)) {
			do {
				
				$GoodUser = CheckOnlineUser
				$index++
			}
			until ($GoodUser -or ($index -ge $RetryCount))
			
			if (!($index -ge $RetryCount)) {
				# Bind to User
				# Check to make Sure user exist
				# Clearing Attributes and Creating Mailbox
				Try {
					Write-Host "Clearing Attributes and Creating Office 365 Mailbox for: " $upn -ForegroundColor Green -BackgroundColor Black
					$RoutingAddress += $user.SamAccountName + "@" + $TenantRoutingDomain.Name
					$ldapCon = "LDAP://" + $user.distinguishedName
					$ADobj = New-Object DirectoryServices.DirectoryEntry $ldapCOn
					$ADobj.Putex(1, "msExchHomeServerName", $null)
					$ADobj.SetInfo()
					Enable-OnPremRemoteMailbox -Identity $upn -RemoteRoutingAddress $RoutingAddress -WarningAction silentlyContinue -ErrorAction Stop
					Invoke-Expression .\StartSync.ps1
					
				}
				Catch { Write-Host "Error Creating Office 365 Mailbox: " $upn -ForegroundColor Red -BackgroundColor Black }
				Finally { }
			}
			
			Else {
				Write-Host "Replication Error or Timeout " -ForegroundColor Red -BackgroundColor Black
				Write-Host " "
				Write-Host " "
			}
			
			
		}
		
		Else {
			Write-Host "User Already Has Office 365 Mailbox: " $upn -ForegroundColor Green -BackgroundColor Black
			Write-Host " "
			Write-Host " "
		}
		
	}
	
	
	Else {
		CLS
		Write-Host""
		Write-Host""
		Write-Host "User Does Not Exist." -ForegroundColor Red -BackgroundColor Black
		Write-Host""
		Write-Host""
	}
}


