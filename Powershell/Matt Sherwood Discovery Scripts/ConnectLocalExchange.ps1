#########################################################################################
# COMPANY: CDW                                                                          #
# NAME:  ConnectLocalExchange.ps1                                                       #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  08/28/2014                                                                     #
#                                                                                       #
# EMAIL: Dean.Sesko@s3.cdw.com                                                          #
#                                                                                       #
# COMMENT:  Script to connec to Local Exchange Server                                   #
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 08/28/2014 Initial Version.                                                       #
# 1.1 10/17/2014 Foramt Cleanup.                                                        #
#                                                                                       #
#########################################################################################
Import-Module ActiveDirectory

#	Setup Shell
$bg = (Get-Host).UI.RawUI
$BG.BackgroundColor = "black"

cls

Write-Host ""
Write-Host ""


# Do Not Change Below this Line

Function GetnewCreds{
	
	Write-Host ""
	$OnPremAdmin = Read-Host -prompt "Please Enter a New On-Premesis Domain Admin Account Ex. Contoso\Administrator"
	$OnPrempassword = read-host -assecurestring -prompt "Password"
	$global:ExchServer = Read-Host -prompt "Please Enter the Name of a Local Exchange 2007/2010/2013 Server "
	$global:HybridServer = Read-Host -prompt "Please Enter the Name of the Hybrid Exchange End Point  "
	$global:OnPremcred = new-object -typename System.Management.Automation.PSCredential -argumentlis $OnPremAdmin, $OnPrempassword
}

Function ConnectLocalService {
Try{
$LocalSession = New-PSSession -Name "ON-Prem-Exchange" -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$ExchServer/PowerShell/" -Authentication kerberos -Credential $OnPremcred -ErrorAction stop
Import-PSSession $LocalSession -Prefix OnPrem -DisableNameChecking -WarningAction silentlyContinue |Out-Null
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
$index =0
do
{
	$index ++
	GetnewCreds
	$Connected = ConnectLocalService}
until ($Connected -or ($index = 5))
Write-Host "Connected to Local Exchange Server......" -ForegroundColor Green -BackgroundColor Black