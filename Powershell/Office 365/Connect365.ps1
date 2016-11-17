#########################################################################################
# COMPANY: CDW                                                                          #
# NAME: Connect365                                                                      #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  08/28/2014                                                                     #
#                                                                                       #
# EMAIL: DeanSesko@planetsesko.com                                                      #
#                                                                                       #
# COMMENT:  Script to connec to Office 365 Administrator Shell and Exchnage Online      #
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 08/28/2014 Initial Version.                                                       #
# 1.1 10/17/2014 Foramt Cleanup and connection detection change.                        #
#                                                                                       #
#########################################################################################
#	Setup Shell
Import-Module exonline
$bg = (Get-Host).UI.RawUI
$BG.BackgroundColor = "black"
cls
Write-Host ""
Write-Host ""


# Do Not Change Below this Line
Function GetnewCreds{

[string]$Global:O365Admin = Read-Host -prompt "Please Enter a New Tenant Admin Account Ex. Admin@tenant.onmicrosoft.com" 
$global:O365password  =read-host -AsSecureString -prompt "Please Enter you Password" 
$global:LiveCred = new-object -typename System.Management.Automation.PSCredential -argumentlist $global:O365Admin,$global:O365password 
}

Function ConnectOnlineServices {
Try{
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell-liveID/ -Credential $LiveCred -Authentication Basic -AllowRedirection  -erroraction silentlycontinue
Import-PSSession $Session 
import-module MSOnline  
Connect-MsolService -Credential $global:LiveCred
return $true
}
Catch{
cls
Write-Host ""
Write-Host ""
Write-Host "Access Denied.  Check your UserName / Password and Try Again" -ForegroundColor Red -BackgroundColor Black
Write-Host ""
Write-Host ""
Write-Host ""
}
Finally{}
}

do {
GetnewCreds
$Connected = ConnectOnlineServices}
until ($Connected)

cls
Write-Host "Connected to Office 365 and Exchange online......" -ForegroundColor Green -BackgroundColor Black