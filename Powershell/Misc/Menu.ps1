####################################################################################################
#                                                                                                  #
#      PowerShell Script to Run all PowerShell Scripts                                             #
#      Author: Brad Stevens                                                                        #
#      Description: Menu Screen for Automating Scripts                                             #
#      Requires: Domain Administrator Credentials and RSAT 2008 R2                                 #
#                                                                                                  #
####################################################################################################

$a = (Get-Host).UI.RawUI
$a.BackgroundColor = "black"
$a.ForegroundColor = "magenta"
$c = $a.BufferSize
$c.Width = 100
$c.Height = 35
$a.BufferSize = $c
$b = $a.WindowSize
$b.Width = 100
$b.Height = 35
$a.WindowSize = $b

CLS


$xAppName    = ‘Menu’
[BOOLEAN]$global:xExitSession=$false
function LoadMenuSystem(){
[INT]$xMenu1=0
[INT]$xMenu2=0
[BOOLEAN]$xValidSelection=$false
while ( $xMenu1 -lt 1 -or $xMenu1 -gt 5 ){
CLS
#… Present the Menu Options
Write-Host “`n`tBrads Ultimate Automation Script`n” -ForegroundColor Magenta
Write-Host “`t`tPlease Select The Automation Category`n” -Fore Cyan
Write-Host “`t`t`t1. Install Microsoft Exchange Server” -Fore Cyan
Write-Host “`t`t`t2. Distribution Group Tasks” -Fore Cyan
Write-Host “`t`t`t3. User Mailbox Tasks” -Fore Cyan
Write-Host "`t`t`t4. Software Installation" -Fore Cyan
Write-Host “`t`t`t5. Quit and exit`n” -Fore Cyan
#… Retrieve the response from the user
[int]$xMenu1 = Read-Host “`tEnter Menu Option Number”
if( $xMenu1 -lt 1 -or $xMenu1 -gt 5 ){
Write-Host “`tPlease select one of the options available.`n” -Fore Red;start-Sleep -Seconds 1
}
}
Switch ($xMenu1){ 
1 {
while ( $xMenu2 -lt 1 -or $xMenu2 -gt 6 ){
CLS
# Present the Menu Options
Write-Host “`n`tBrads Ultimate Automation Script`n” -Fore Magenta
Write-Host “`t`tPlease select the version of Exchange Server you would like to install`n” -Fore Green
Write-Host “`t`t`t1. Exchange Server 2013 CU9” -Fore Green
Write-Host “`t`t`t2. Exchange Server 2013 CU10” -Fore Green
Write-Host “`t`t`t3. Exchange Server 2016 RTM” -Fore Green
Write-Host “`t`t`t4. Go to Main Menu`n” -Fore Green
[int]$xMenu2 = Read-Host “`tEnter Menu Option Number”
if( $xMenu1 -lt 1 -or $xMenu1 -gt 4 ){
Write-Host “`tPlease select one of the options available.`n” -Fore Red;start-Sleep -Seconds 1
}
}
Switch ($xMenu2){
1{
 #####################################################################################
   #                                                                                   #
   #                                 Functions                                         #
   #                                                                                   #
   #####################################################################################


    Function Package-Download($Package_Url, $Package_File) {
        $Package_Stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Package_File, Create
        $Package_URI = New-Object "System.Uri" "$Package_Url"
        $Package_Request = [System.Net.HttpWebRequest]::Create($Package_URI)
        # 30 Second Timeout
        $Package_Request.set_Timeout(30000)
        $Package_Response = $Package_Request.GetResponse()
        $Package_Total = [System.Math]::Floor($Package_Response.get_ContentLength()/1024)
        $Package_ResponseStream = $Package_Response.GetResponseStream()
        $Package_Buffer = new-object byte[] 10KB
        $Package_Count = $Package_ResponseStream.Read($Package_Buffer,0,$Package_Buffer.length)
        $Package_DownloadedBytes = $Package_Count
        while ($Package_Count -gt 0) {
            $Package_Stream.Write($Package_Buffer, 0, $Package_Count)
            $Package_Count = $Package_ResponseStream.Read($Package_Buffer,0,$Package_Buffer.length)
            $Package_DownloadedBytes = $Package_DownloadedBytes + $Package_Count
            Write-Progress -activity "Downloading file '$($Package_Url.split('/') | Select -Last 1)'" -status "Downloaded ($([System.Math]::Floor($Package_DownloadedBytes/1024))K of $($Package_Total)K): " -PercentComplete ((([System.Math]::Floor($Package_DownloadedBytes/1024)) / $Package_Total)  * 100)
         }
        Write-Progress -activity "Finished downloading file '$($Package_Url.split('/') | Select -Last 1)'"
        $Package_Stream.Flush()
        $Package_Stream.Close()
        $Package_Stream.Dispose()
        $Package_ResponseStream.Dispose()
    }

    Function Disable-UAC {
        Write-Output "Disabling User Account Control..."
        New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 0 -Force | out-null
        Write-Output "Disabled User Account Control!"
    }

    Function Disable-IEESC {
        Write-Output "Disabling IE Enhanced Security Configuration..."
        $AdminKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"
        $UserKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}"
        Set-ItemProperty -Path $AdminKey -Name "IsInstalled" -Value 0
        Set-ItemProperty -Path $UserKey -Name "IsInstalled" -Value 0
        Stop-Process -Name Explorer -Force
        Write-Output "IE Enhanced Security Configuration Disabled!"
    }
    
    Function Disable-OpenFileSecurityWarning {
        Write-Output "Disabling File Security Warning dialog..."
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations" -ErrorAction SilentlyContinue |out-null
        New-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations" -name "LowRiskFileTypes" -value ".exe;.msp;.msu" -ErrorAction SilentlyContinue |out-null
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments" -ErrorAction SilentlyContinue |out-null
        New-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments" -name "SaveZoneInformation" -value 1 -ErrorAction SilentlyContinue |out-null
        Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations" -Name "LowRiskFileTypes" -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments" -Name "SaveZoneInformation" -ErrorAction SilentlyContinue
        Write-Output "File Security Warning Disabled!"
    }


   #####################################################################################
   #                                                                                   #
   #                                 Main Script                                       #
   #                                                                                   #
   #####################################################################################


If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
    $Arguments = "& '" + $MyInvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
}

# Prompt for Active Directory Domain Credentials
    $Credentials = $host.ui.PromptForCredential("Domain Administrator Credentials", "Please enter your domain user name and password in the following format: domain\username.", "", "NetBiosUserName")
    $Domain = $Credentials.GetNetworkCredential().Domain 
    $PlainUsername = $Credentials.GetNetworkCredential().UserName
    $PlainPassword = $Credentials.GetNetworkCredential().Password
# Loads System.Reflection.Assembly to prompt user for Exchange Organization Name
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    $OrganizationName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the Exchange Organization Name.", "Exchange Organization", "")
# RegEx to remove spaces in Organization Name 
    $OrganizationName = $OrganizationName -replace '\s+', '' 
# Registry Changes to allow Auto Logon with specified Active Directory Domain Credentials
    $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"  
    Set-ItemProperty $RegPath "AutoAdminLogon" -Value "1" -type String  
    Set-ItemProperty $RegPath "DefaultUsername" -Value "$Domain\$PlainUsername" -type String  
    Set-ItemProperty $RegPath "DefaultPassword" -Value "$PlainPassword" -type String
    Set-ItemProperty -Path HKLM:\SOFTWARE\Classes\Microsoft.PowerShellScript.1\Shell\Open\Command -Name "(Default)" -Value '"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" "-file" "%1"' -Type String -Force
    Set-ItemProperty -Path HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe -Name "QuickEdit" -Value '0' -Type DWord -Force
    New-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage -Name "OpenAtLogon" -Value '0' -Type DWord -Force
    New-ItemProperty -Path HKCU:\Software\Microsoft\ServerManager -Name "DoNotOpenServerManagerAtLogon" -Value "1" -Type DWord -Force
# Checks if C:\Exchange Resources already exists. If it does not, it creates the Exchange Resources Directory
    $Directory = Test-Path "C:\Exchange Resources"
    if($Directory -eq $false) {
        New-Item "C:\Exchange Resources" -type directory
    }
    New-Item "C:\Exchange Resources\OrganizationName.txt" -ItemType File
    Set-Content -Path "C:\Exchange Resources\OrganizationName.txt" -Value "$OrganizationName"

# Clears Host and Implements 4 new lines to make write-host prompts visible
    Clear-Host
    Write-Host `n
    Write-Host `n
    Write-Host `n
    Write-Host `n
    Disable-UAC
    Disable-OpenFileSecurityWarning
    Disable-IEESC
    
    Write-Host "Downloading Installation Packages..." `n

# Download request for Exchange 2013 CU9
    Package-Download "http://download.microsoft.com/download/C/6/8/C6899C99-F933-4181-9692-17A5BB7F1A4B/Exchange2013-x64-cu9.exe" "C:\Exchange Resources\Exchange2013-x64-cu9.exe" 
    Write-Host "Finished downloading Exchange2013-x64-cu9.exe"
    Write-Host "Downloading Filter Pack SP1..."

# Download request for 2010 Filter Pack SP1
    Package-Download "http://download.microsoft.com/download/A/A/3/AA345161-18B8-45AE-8DC8-DA6387264CB9/filterpack2010sp1-kb2460041-x64-fullfile-en-us.exe" "C:\Exchange Resources\filterpack2010sp1-kb2460041-x64-fullfile-en-us.exe"
    Write-Host "Finished downloading filterpack2010sp1-kb2460041-x64-fullfile-en-us.exe"
    Write-Host "Downloading UCMA Runtime..."

# Download request for UcmaRuntimeSetup
    Package-Download "http://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe" "C:\Exchange Resources\UcmaRuntimeSetup.exe"
    Write-Host "Finished downloading UcmaRuntimeSetup.exe"
    Write-Host "Downloading .NET Framework 4.5.2..."

# Download request for .NET Framework 4.5.2
    Package-Download "http://download.microsoft.com/download/E/2/1/E21644B5-2DF2-47C2-91BD-63C560427900/NDP452-KB2901907-x86-x64-AllOS-ENU.exe" "C:\Exchange Resources\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
    Write-Host "Finished downloading .NET Framework 4.5.2!"
    Write-Host "Installing Windows Feature Prerequisites..."

# Installation of Windows Feature Prerequisites
    Install-WindowsFeature RSAT-ADDS
    Install-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
    Write-Host "Finished installing Windows Feature Prerequisites!"
# Specifys Extraction Folder for Exchange Setup
    $Targetfolder="C:\Exchange_2013_CU9_Setup"
# Checks, then Extracts Exchange to Extraction Folder
    $Extract = Test-Path "C:\Exchange_2013_CU9_Setup"
    if ($Extract -eq $false) {
        Start-Process -Filepath "C:\Exchange Resources\Exchange2013-x64-cu9.exe" -ArgumentList "/extract:$Targetfolder /u" -Wait
    }
# Specifies Startup Location for Install Script
    $Path = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
# Tests Startup Location then creates InstallExchange.ps1 file
    $TestPath = Test-Path "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
    if ( $TestPath -eq $false ) {
        New-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1" -type file
    }
    else {
        Write-Host "Script already Exists in $env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\" `n
        Write-Host "Resuming Script..." `n
    }
 # Sets Content for InstallExchange.ps1 script upon 1st reboot to run on startup
    Set-Content $Path {
            If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
            $Arguments = "& '" + $MyInvocation.mycommand.definition + "'"
            Start-Process powershell -Verb runAs -ArgumentList $arguments
            Break
        } 
        Remove-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
        Start-Process -FilePath "C:\Exchange Resources\NDP452-KB2901907-x86-x64-AllOS-ENU.exe" -ArgumentList "/passive /norestart" -wait                
        Start-Process -Filepath "C:\Exchange Resources\UcmaRuntimeSetup.exe" -ArgumentList "/passive /norestart" -wait
        Start-Process -Filepath "C:\Exchange Resources\filterpack2010sp1-kb2460041-x64-fullfile-en-us.exe" -ArgumentList "/passive /norestart" -wait
        $Path = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
        $TestPath = Test-Path "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
        if( $TestPath -eq $false ) {
            New-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1" -type file
        }
        else {
            Write-Host "Script already Exists in $env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\" `n
            Write-Host "Resuming Script..." `n
        }
 # Sets Content for InstallExchange.ps1 script upon 2nd reboot to run on startup
        Set-Content $Path {
            If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
                $Arguments = "& '" + $MyInvocation.mycommand.definition + "'"
                Start-Process powershell -Verb runAs -ArgumentList $arguments
                Break
            }        
            $Organization = Get-Content -Path "C:\Exchange Resources\OrganizationName.txt" 
            Write-Host "Server is now ready to install Exchange Server 2013 CU9..."
            Function Enable-OpenFileSecurityWarning {
            Write-Output "Enabling File Security Warning dialog..."
            Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations" -Name "LowRiskFileTypes" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments" -Name "SaveZoneInformation" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations" -Name "LowRiskFileTypes" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments" -Name "SaveZoneInformation" -ErrorAction SilentlyContinue
            Write-Output "Finished!"
            }
            Function Enable-IEESC {
            Write-Verbose "Enabling IE Enhanced Security Configuration..."
            $AdminKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"
            $UserKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}"
            Set-ItemProperty -Path $AdminKey -Name "IsInstalled" -Value 1
            Set-ItemProperty -Path $UserKey -Name "IsInstalled" -Value 1
            Stop-Process -Name Explorer
            Write-Output "Finished!"
            }
            Function Enable-UAC {
            Write-Output "Enabling User Account Control..."
            New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 1 -ErrorAction SilentlyContinue| out-null
            Write-Output "Finished!"
            }
            C:\Exchange_2013_CU9_Setup\setup.exe /PrepareSchema /IAcceptExchangeServerLicenseTerms
            C:\Exchange_2013_CU9_Setup\setup.exe /PrepareAD /OrganizationName:homelabnet /IAcceptExchangeServerLicenseTerms
            C:\Exchange_2013_CU9_Setup\setup.exe /PrepareAllDomains /IAcceptExchangeServerLicenseTerms                       
            C:\Exchange_2013_CU9_Setup\Setup.exe /mode:Install /role:ClientAccess,Mailbox /OrganizationName:"$Organization" /IAcceptExchangeServerLicenseTerms
 # Changes file association for .ps1 to default (notepad.exe)
            Set-ItemProperty -Path HKLM:\SOFTWARE\Classes\Microsoft.PowerShellScript.1\Shell\Open\Command -Name "(Default)" -Value '"C:\Windows\System32\notepad.exe" "%1"' -Type String -Force
 # Changes QuickEdit to default registry setting
            Set-ItemProperty -Path HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe -Name "QuickEdit" -Value '1' -Type DWord -Force
 # Changes DoNotOpenServerManagerAtLogon to default registry setting
            New-ItemProperty -Path HKCU:\Software\Microsoft\ServerManager -Name "DoNotOpenServerManagerAtLogon" -Value "0" -Type DWord -Force
 # Removes InstallExchange.ps1 script from startup folder
            Remove-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
 # Removes leftover text file for Organization Name
            Remove-Item "C:\Exchange Resources\OrganizationName.txt"
 # Re-enables UAC
            Enable-UAC
 # Re-enables IEESC
            Enable-IEESC
 # Re-enables OpenFileSecurityWarning
            Enable-OpenFileSecurityWarning
 # Changes AutoAdminLogon to default  registry setting
            $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"  
            Set-ItemProperty $RegPath "AutoAdminLogon" -Value "0" -type String
            Restart-Computer
          }
    Restart-Computer
    }
Restart-Computer

}
2{

$a = (Get-Host).UI.RawUI
$a.BackgroundColor = "black"
$a.ForegroundColor = "magenta"

<#
    .TITLE
    Brad's Ultimate Exchange Server 2013 CU9 Automation Script
    
    .SYNOPSIS
    Installs Exchange Server 2013 CU9 Client Access and Mailbox Roles and all required prerequisites.
    
    Adapted from:
    - http://blogs.msdn.com/b/jasonn/archive/2013/06/11/8594493.aspx;
    - http://eightwone.com/2013/02/18/exchange-2013-unattended-installation-script


    Brad C. Stevens
    brad@bradcstevens.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 1.0, September 8th, 2015

    .DESCRIPTION
    This script will install Exchange 2013 CU9 prerequisites, creates the Exchange 
    organization (prepares Active Directory) and installs Exchange Server. 


    .LINK
    - http://bradcstevens.com
    - http://helpdesksage.blogspot.com

    .NOTES
    Requirements:
    - Windows Server 2012 R2;
    - Domain-joined system;
    - Domain Account with Domain Admins, Enterprise Admins, and Schema Admins membership

    Revision History
    ------------------------------------------------------------------------------------
    1.0 Initial community release
    1.1 Added automatic recognition and start of script with runas "Administrator"
    1.2 Added check for domain membership of the server that is to receive the 
        installation of Exchange. 
        Updated the script to include CU10 instead of CU9
        Updated script to include 2010 Filter Pack SP2
        Added check for already installed prerequisite packages

#>


   #####################################################################################
   #                                                                                   #
   #                                 Functions                                         #
   #                                                                                   #
   #####################################################################################


    Function Package-Download($Package_Url, $Package_File) {
        $Package_Stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Package_File, Create
        $Package_URI = New-Object "System.Uri" "$Package_Url"
        $Package_Request = [System.Net.HttpWebRequest]::Create($Package_URI)
        # 30 Second Timeout
        $Package_Request.set_Timeout(30000)
        $Package_Response = $Package_Request.GetResponse()
        $Package_Total = [System.Math]::Floor($Package_Response.get_ContentLength()/1024)
        $Package_ResponseStream = $Package_Response.GetResponseStream()
        $Package_Buffer = new-object byte[] 10KB
        $Package_Count = $Package_ResponseStream.Read($Package_Buffer,0,$Package_Buffer.length)
        $Package_DownloadedBytes = $Package_Count
        while ($Package_Count -gt 0) {
            $Package_Stream.Write($Package_Buffer, 0, $Package_Count)
            $Package_Count = $Package_ResponseStream.Read($Package_Buffer,0,$Package_Buffer.length)
            $Package_DownloadedBytes = $Package_DownloadedBytes + $Package_Count
            Write-Progress -activity "Downloading file '$($Package_Url.split('/') | Select -Last 1)'" -status "Downloaded ($([System.Math]::Floor($Package_DownloadedBytes/1024))K of $($Package_Total)K): " -PercentComplete ((([System.Math]::Floor($Package_DownloadedBytes/1024)) / $Package_Total)  * 100)
         }
        Write-Progress -activity "Finished downloading file '$($Package_Url.split('/') | Select -Last 1)'"
        $Package_Stream.Flush()
        $Package_Stream.Close()
        $Package_Stream.Dispose()
        $Package_ResponseStream.Dispose()
    }

    Function Disable-UAC {
        Write-Output "Disabling User Account Control..."
        New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 0 -Force | out-null
        Write-Output "Disabled User Account Control!"
    }

    Function Disable-IEESC {
        Write-Output "Disabling IE Enhanced Security Configuration..."
        $AdminKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"
        $UserKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}"
        Set-ItemProperty -Path $AdminKey -Name "IsInstalled" -Value 0
        Set-ItemProperty -Path $UserKey -Name "IsInstalled" -Value 0
        Stop-Process -Name Explorer -Force
        Write-Output "IE Enhanced Security Configuration Disabled!"
    }
    
    Function Disable-OpenFileSecurityWarning {
        Write-Output "Disabling File Security Warning dialog..."
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations" -ErrorAction SilentlyContinue |out-null
        New-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations" -name "LowRiskFileTypes" -value ".exe;.msp;.msu" -ErrorAction SilentlyContinue |out-null
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments" -ErrorAction SilentlyContinue |out-null
        New-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments" -name "SaveZoneInformation" -value 1 -ErrorAction SilentlyContinue |out-null
        Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations" -Name "LowRiskFileTypes" -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments" -Name "SaveZoneInformation" -ErrorAction SilentlyContinue
        Write-Output "File Security Warning Disabled!"
    }


   #####################################################################################
   #                                                                                   #
   #                                 Main Script                                       #
   #                                                                                   #
   #####################################################################################


    If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
    $Arguments = "& '" + $MyInvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
    }
    Write-Host "Checking server domain membership....." `n
    If((gwmi win32_computersystem).partofdomain -eq $true) {
    Write-Host "Server is Domain Joined! Resuming Script...."
    }   
    Else {
    Write-Host -fore red "Domain Membership Required. Join this server to a domain then run the script."
    Pause
    Exit
    }
# Prompt for Active Directory Domain Credentials
    $Credentials = $host.ui.PromptForCredential("Domain Administrator Credentials", "Please enter your domain user name and password in the following format: domain\username.", "", "NetBiosUserName")
    $Domain = $Credentials.GetNetworkCredential().Domain 
    $PlainUsername = $Credentials.GetNetworkCredential().UserName
    $PlainPassword = $Credentials.GetNetworkCredential().Password
# Loads System.Reflection.Assembly to prompt user for Exchange Organization Name
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    $OrganizationName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the Exchange Organization Name.", "Exchange Organization", "")
# RegEx to remove spaces in Organization Name 
    $OrganizationName = $OrganizationName -replace '\s+', '' 
# Registry Changes to allow Auto Logon with specified Active Directory Domain Credentials
    $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"  
    Set-ItemProperty $RegPath "AutoAdminLogon" -Value "1" -type String  
    Set-ItemProperty $RegPath "DefaultUsername" -Value "$Domain\$PlainUsername" -type String  
    Set-ItemProperty $RegPath "DefaultPassword" -Value "$PlainPassword" -type String
    Set-ItemProperty -Path HKLM:\SOFTWARE\Classes\Microsoft.PowerShellScript.1\Shell\Open\Command -Name "(Default)" -Value '"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" "-file" "%1"' -Type String -Force
    Set-ItemProperty -Path HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe -Name "QuickEdit" -Value '0' -Type DWord -Force
    New-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage -Name "OpenAtLogon" -Value '0' -Type DWord -Force
    New-ItemProperty -Path HKCU:\Software\Microsoft\ServerManager -Name "DoNotOpenServerManagerAtLogon" -Value "1" -Type DWord -Force
# Checks if C:\Exchange Resources already exists. If it does not, it creates the Exchange Resources Directory
    $Directory = Test-Path "C:\Exchange Resources"
    if($Directory -eq $false) {
        New-Item "C:\Exchange Resources" -type directory
    }
    $OrganizationFilePath = Test-Path "C:\Exchange Resources\OrganizationName.txt"
    if($OrganizationFilePath -eq $false) {
        New-Item "C:\Exchange Resources\OrganizationName.txt" -ItemType File
        Set-Content -Path "C:\Exchange Resources\OrganizationName.txt" -Value "$OrganizationName"
    }

# Clears Host and Implements 4 new lines to make write-host prompts visible
    Clear-Host
    Write-Host `n
    Write-Host `n
    Write-Host `n
    Write-Host `n
    Disable-UAC
    Disable-OpenFileSecurityWarning
    Disable-IEESC
    
    Write-Host "Downloading Installation Packages..." `n

# Download request for Exchange 2013 CU9
    $Exchange2013 = Test-Path "C:\Exchange Resources\Exchange2013-x64-cu10.exe"
    If($Exchange2013 -eq $false) { 
    Package-Download "http://download.microsoft.com/download/1/D/1/1D15B640-E2BB-4184-BFC5-83BC26ADD689/Exchange2013-x64-cu10.exe" "C:\Exchange Resources\Exchange2013-x64-cu10.exe" 
    Write-Host "Finished downloading Exchange2013-x64-cu10.exe"
    Write-Host "Downloading Filter Pack SP1..."
    }
# Download request for 2010 Filter Pack SP2
    $FilterPackSp2 = Test-Path "C:\Exchange Resources\filterpack2010sp2-x64-fullfile-en-us.exe"
    If($FilterPackSp2 -eq $false) {
    Package-Download "http://download.microsoft.com/download/D/C/A/DCA32A51-6954-4814-8838-422BD3F508F8/filterpacksp2010-kb2687447-fullfile-x64-en-us.exe" "C:\Exchange Resources\filterpack2010sp2-x64-fullfile-en-us.exe"
    Write-Host "Finished downloading filterpack2010sp2-x64-fullfile-en-us.exe"
    Write-Host "Downloading UCMA Runtime..."
    }
# Download request for UcmaRuntimeSetup
    $UCMARuntimeSetup = Test-Path "C:\Exchange Resources\UcmaRuntimeSetup.exe"
    If($UCMARuntimeSetup -eq $false) {
    Package-Download "http://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe" "C:\Exchange Resources\UcmaRuntimeSetup.exe"
    Write-Host "Finished downloading UcmaRuntimeSetup.exe"
    Write-Host "Downloading .NET Framework 4.5.2..."
    }
# Download request for .NET Framework 4.5.2
    $NDP452 = Test-Path "C:\Exchange Resources\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
    If($NDP452 -eq $false) {
    Package-Download "http://download.microsoft.com/download/E/2/1/E21644B5-2DF2-47C2-91BD-63C560427900/NDP452-KB2901907-x86-x64-AllOS-ENU.exe" "C:\Exchange Resources\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
    Write-Host "Finished downloading .NET Framework 4.5.2!"
    Write-Host "Installing Windows Feature Prerequisites..."
    }
# Installation of Windows Feature Prerequisites
    Install-WindowsFeature RSAT-ADDS
    Install-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
    Write-Host "Finished installing Windows Feature Prerequisites!"
# Specifys Extraction Folder for Exchange Setup
    $Targetfolder="C:\Exchange_2013_CU10_Setup"
# Checks, then Extracts Exchange to Extraction Folder
    $Extract = Test-Path "C:\Exchange_2013_CU10_Setup"
    if ($Extract -eq $false) {
        Start-Process -Filepath "C:\Exchange Resources\Exchange2013-x64-cu10.exe" -ArgumentList "/extract:$Targetfolder /u" -Wait
    }
# Specifies Startup Location for Install Script
    $Path = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
# Tests Startup Location then creates InstallExchange.ps1 file
    $TestPath = Test-Path "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
    if ( $TestPath -eq $false ) {
        New-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1" -type file
    }
    else {
        Write-Host "Script already Exists in $env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\" `n
        Write-Host "Resuming Script..." `n
    }
 # Sets Content for InstallExchange.ps1 script upon 1st reboot to run on startup
    Set-Content $Path {
            If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
            $Arguments = "& '" + $MyInvocation.mycommand.definition + "'"
            Start-Process powershell -Verb runAs -ArgumentList $arguments
            Break
        } 
        Remove-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
        Start-Process -FilePath "C:\Exchange Resources\NDP452-KB2901907-x86-x64-AllOS-ENU.exe" -ArgumentList "/passive /norestart" -wait                
        Start-Process -Filepath "C:\Exchange Resources\UcmaRuntimeSetup.exe" -ArgumentList "/passive /norestart" -wait
        Start-Process -Filepath "C:\Exchange Resources\filterpack2010sp2-x64-fullfile-en-us.exe" -ArgumentList "/passive /norestart" -wait
        $Path = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
        $TestPath = Test-Path "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
        if( $TestPath -eq $false ) {
            New-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1" -type file
        }
        else {
            Write-Host "Script already Exists in $env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\" `n
            Write-Host "Resuming Script..." `n
        }
 # Sets Content for InstallExchange.ps1 script upon 2nd reboot to run on startup
        Set-Content $Path {
            If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
                $Arguments = "& '" + $MyInvocation.mycommand.definition + "'"
                Start-Process powershell -Verb runAs -ArgumentList $arguments
                Break
            }        
            $Organization = Get-Content -Path "C:\Exchange Resources\OrganizationName.txt" 
            Write-Host "Server is now ready to install Exchange Server 2013 CU10..."
            Function Enable-OpenFileSecurityWarning {
            Write-Output "Enabling File Security Warning dialog..."
            Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations" -Name "LowRiskFileTypes" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments" -Name "SaveZoneInformation" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations" -Name "LowRiskFileTypes" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments" -Name "SaveZoneInformation" -ErrorAction SilentlyContinue
            Write-Output "Finished!"
            }
            Function Enable-IEESC {
            Write-Verbose "Enabling IE Enhanced Security Configuration..."
            $AdminKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"
            $UserKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}"
            Set-ItemProperty -Path $AdminKey -Name "IsInstalled" -Value 1
            Set-ItemProperty -Path $UserKey -Name "IsInstalled" -Value 1
            Stop-Process -Name Explorer
            Write-Output "Finished!"
            }
            Function Enable-UAC {
            Write-Output "Enabling User Account Control..."
            New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 1 -ErrorAction SilentlyContinue| out-null
            Write-Output "Finished!"
            }
            C:\Exchange_2013_CU10_Setup\setup.exe /PrepareSchema /IAcceptExchangeServerLicenseTerms
            C:\Exchange_2013_CU10_Setup\setup.exe /PrepareAD /OrganizationName:"$Organization" /IAcceptExchangeServerLicenseTerms
            C:\Exchange_2013_CU10_Setup\setup.exe /PrepareAllDomains /IAcceptExchangeServerLicenseTerms                       
            C:\Exchange_2013_CU10_Setup\Setup.exe /mode:Install /role:ClientAccess,Mailbox /OrganizationName:"$Organization" /IAcceptExchangeServerLicenseTerms
 # Changes file association for .ps1 to default (notepad.exe)
            Set-ItemProperty -Path HKLM:\SOFTWARE\Classes\Microsoft.PowerShellScript.1\Shell\Open\Command -Name "(Default)" -Value '"C:\Windows\System32\notepad.exe" "%1"' -Type String -Force
 # Changes QuickEdit to default registry setting
            Set-ItemProperty -Path HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe -Name "QuickEdit" -Value '1' -Type DWord -Force
 # Changes DoNotOpenServerManagerAtLogon to default registry setting
            New-ItemProperty -Path HKCU:\Software\Microsoft\ServerManager -Name "DoNotOpenServerManagerAtLogon" -Value "0" -Type DWord -Force
 # Removes InstallExchange.ps1 script from startup folder
            Remove-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
 # Removes leftover text file for Organization Name
            Remove-Item "C:\Exchange Resources\OrganizationName.txt"
 # Re-enables UAC
            Enable-UAC
 # Re-enables IEESC
            Enable-IEESC
 # Re-enables OpenFileSecurityWarning
            Enable-OpenFileSecurityWarning
 # Changes AutoAdminLogon to default  registry setting
            $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"  
            Set-ItemProperty $RegPath "AutoAdminLogon" -Value "0" -type String
            Restart-Computer
          }
    Restart-Computer
    }
Restart-Computer

}
3{ 



$a = (Get-Host).UI.RawUI
$a.BackgroundColor = "black"
$a.ForegroundColor = "magenta"


<#
    .TITLE
    Brad's Ultimate Exchange Server 2016 Automation Script
    
    .SYNOPSIS
    Installs Exchange Server 2016 Mailbox Role and all required prerequisites.
    
    Adapted from:
    - http://blogs.msdn.com/b/jasonn/archive/2013/06/11/8594493.aspx;
    - http://eightwone.com/2013/02/18/exchange-2013-unattended-installation-script


    Brad C. Stevens
    brad@bradcstevens.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 1.0, November 1, 2015

    .DESCRIPTION
    This script will install Exchange 2013 CU9 prerequisites, creates the Exchange 
    organization (prepares Active Directory) and installs Exchange Server. 


    .LINK
    - http://bradcstevens.com
    - http://helpdesksage.blogspot.com

    .NOTES
    Requirements:
    - Windows Server 2012 R2;
    - Domain-joined system;
    - Domain Account with Domain Admins, Enterprise Admins, and Schema Admins membership

    Revision History
    ------------------------------------------------------------------------------------
    1.0 Initial community release
    1.1 Added automatic recognition and start of script with runas "Administrator"
    1.2 Added check for domain membership of the server that is to receive the 
        installation of Exchange. 
        Updated the script to include CU10 instead of CU9
        Updated script to include 2010 Filter Pack SP2
        Added check for already installed prerequisite packages

#>


   #####################################################################################
   #                                                                                   #
   #                                 Functions                                         #
   #                                                                                   #
   #####################################################################################


    Function Package-Download($Package_Url, $Package_File) {
        $Package_Stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Package_File, Create
        $Package_URI = New-Object "System.Uri" "$Package_Url"
        $Package_Request = [System.Net.HttpWebRequest]::Create($Package_URI)
        # 30 Second Timeout
        $Package_Request.set_Timeout(30000)
        $Package_Response = $Package_Request.GetResponse()
        $Package_Total = [System.Math]::Floor($Package_Response.get_ContentLength()/1024)
        $Package_ResponseStream = $Package_Response.GetResponseStream()
        $Package_Buffer = new-object byte[] 1000KB
        $Package_Count = $Package_ResponseStream.Read($Package_Buffer,0,$Package_Buffer.length)
        $Package_DownloadedBytes = $Package_Count
        while ($Package_Count -gt 0) {
            $Package_Stream.Write($Package_Buffer, 0, $Package_Count)
            $Package_Count = $Package_ResponseStream.Read($Package_Buffer,0,$Package_Buffer.length)
            $Package_DownloadedBytes = $Package_DownloadedBytes + $Package_Count
            Write-Progress -activity "Downloading file '$($Package_Url.split('/') | Select -Last 1)'" -status "Downloaded ($([System.Math]::Floor($Package_DownloadedBytes/1024))K of $($Package_Total)K): " -PercentComplete ((([System.Math]::Floor($Package_DownloadedBytes/1024)) / $Package_Total)  * 100)
         }
        Write-Progress -activity "Finished downloading file '$($Package_Url.split('/') | Select -Last 1)'"
        $Package_Stream.Flush()
        $Package_Stream.Close()
        $Package_Stream.Dispose()
        $Package_ResponseStream.Dispose()
    }

    Function Disable-UAC {
        Write-Output "Disabling User Account Control..."
        New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 0 -Force | out-null
        Write-Output "Disabled User Account Control!"
    }

    Function Disable-IEESC {
        Write-Output "Disabling IE Enhanced Security Configuration..."
        $AdminKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"
        $UserKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}"
        Set-ItemProperty -Path $AdminKey -Name "IsInstalled" -Value 0
        Set-ItemProperty -Path $UserKey -Name "IsInstalled" -Value 0
        Stop-Process -Name Explorer -Force
        Write-Output "IE Enhanced Security Configuration Disabled!"
    }
    
    Function Disable-OpenFileSecurityWarning {
        Write-Output "Disabling File Security Warning dialog..."
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations" -ErrorAction SilentlyContinue |out-null
        New-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations" -name "LowRiskFileTypes" -value ".exe;.msp;.msu" -ErrorAction SilentlyContinue |out-null
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments" -ErrorAction SilentlyContinue |out-null
        New-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments" -name "SaveZoneInformation" -value 1 -ErrorAction SilentlyContinue |out-null
        Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations" -Name "LowRiskFileTypes" -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments" -Name "SaveZoneInformation" -ErrorAction SilentlyContinue
        Write-Output "File Security Warning Disabled!"
    }


   #####################################################################################
   #                                                                                   #
   #                                 Main Script                                       #
   #                                                                                   #
   #####################################################################################


    If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
    $Arguments = "& '" + $MyInvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
    }
    Write-Host "Checking server domain membership....." `n
    If((gwmi win32_computersystem).partofdomain -eq $true) {
    Write-Host "Server is Domain Joined! Resuming Script...."
    }   
    Else {
    Write-Host -fore red "Domain Membership Required. Join this server to a domain then run the script."
    Pause
    Exit
    }
# Prompt for Active Directory Domain Credentials
    $Credentials = $host.ui.PromptForCredential("Domain Administrator Credentials", "Please enter your domain user name and password in the following format: domain\username.", "", "NetBiosUserName")
    $Domain = $Credentials.GetNetworkCredential().Domain 
    $PlainUsername = $Credentials.GetNetworkCredential().UserName
    $PlainPassword = $Credentials.GetNetworkCredential().Password
# Loads System.Reflection.Assembly to prompt user for Exchange Organization Name
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    $OrganizationName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the Exchange Organization Name.", "Exchange Organization", "")
# RegEx to remove spaces in Organization Name 
    $OrganizationName = $OrganizationName -replace '\s+', '' 
# Registry Changes to allow Auto Logon with specified Active Directory Domain Credentials
    $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"  
    Set-ItemProperty $RegPath "AutoAdminLogon" -Value "1" -type String  
    Set-ItemProperty $RegPath "DefaultUsername" -Value "$Domain\$PlainUsername" -type String  
    Set-ItemProperty $RegPath "DefaultPassword" -Value "$PlainPassword" -type String
    Set-ItemProperty -Path HKLM:\SOFTWARE\Classes\Microsoft.PowerShellScript.1\Shell\Open\Command -Name "(Default)" -Value '"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" "-file" "%1"' -Type String -Force
    Set-ItemProperty -Path HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe -Name "QuickEdit" -Value '0' -Type DWord -Force
    New-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage -Name "OpenAtLogon" -Value '0' -Type DWord -Force
    New-ItemProperty -Path HKCU:\Software\Microsoft\ServerManager -Name "DoNotOpenServerManagerAtLogon" -Value "1" -Type DWord -Force
# Checks if C:\Exchange Resources already exists. If it does not, it creates the Exchange Resources Directory
    $Directory = Test-Path "C:\Exchange Resources"
    if($Directory -eq $false) {
        New-Item "C:\Exchange Resources" -type directory
    }
    $OrganizationFilePath = Test-Path "C:\Exchange Resources\OrganizationName.txt"
    if($OrganizationFilePath -eq $false) {
        New-Item "C:\Exchange Resources\OrganizationName.txt" -ItemType File
        Set-Content -Path "C:\Exchange Resources\OrganizationName.txt" -Value "$OrganizationName"
    }

# Clears Host and Implements 4 new lines to make write-host prompts visible
    Clear-Host
    Write-Host `n
    Write-Host `n
    Write-Host `n
    Write-Host `n
    Disable-UAC
    Disable-OpenFileSecurityWarning
    Disable-IEESC
    
    Write-Host "Downloading Installation Packages..." `n

# Download request for Exchange 2016
    $Exchange2016 = Test-Path "C:\Exchange Resources\Exchange2016-x64-cu10.exe"
    If($Exchange2016 -eq $false) { 
    Package-Download "https://download.microsoft.com/download/3/9/B/39B8DDA8-509C-4B9E-BCE9-4CD8CDC9A7DA/Exchange2016-x64.exe" "C:\Exchange Resources\Exchange2016-x64" 
    Write-Host "Finished downloading Exchange2016-x64.exe"
    Write-Host "Downloading UCMARuntime..."
    }
 Download request for .NET Framework 4.5.2
    $NDP452 = Test-Path "C:\Exchange Resources\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
    If($NDP452 -eq $false) {
    Package-Download "http://download.microsoft.com/download/E/2/1/E21644B5-2DF2-47C2-91BD-63C560427900/NDP452-KB2901907-x86-x64-AllOS-ENU.exe" "C:\Exchange Resources\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
    Write-Host "Finished downloading .NET Framework 4.5.2!"
    Write-Host "Installing Windows Feature Prerequisites..."
    }
# Download request for UcmaRuntimeSetup
    $UCMARuntimeSetup = Test-Path "C:\Exchange Resources\UcmaRuntimeSetup.exe"
    If($UCMARuntimeSetup -eq $false) {
    Package-Download "http://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe" "C:\Exchange Resources\UcmaRuntimeSetup.exe"
    Write-Host "Finished downloading UcmaRuntimeSetup.exe"
    Write-Host "Installing Windows Features..."
    }
    Install-WindowsFeature RSAT-ADDS
    Install-WindowsFeature Install-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
    Write-Host "Finished installing Windows Feature Prerequisites!"
# Specifys Extraction Folder for Exchange Setup
    $Targetfolder="C:\Exchange_2016_Setup"
# Checks, then Extracts Exchange to Extraction Folder
    $Extract = Test-Path "C:\Exchange_2016_Setup"
    if ($Extract -eq $false) {
        Start-Process -Filepath "C:\Exchange Resources\Exchange2016-x64.exe" -ArgumentList "/extract:$Targetfolder /u" -Wait
    }
# Specifies Startup Location for Install Script
    $Path = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
# Tests Startup Location then creates InstallExchange.ps1 file
    $TestPath = Test-Path "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
    if ( $TestPath -eq $false ) {
        New-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1" -type file
    }
    else {
        Write-Host "Script already Exists in $env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\" `n
        Write-Host "Resuming Script..." `n
    }
 # Sets Content for InstallExchange.ps1 script upon 1st reboot to run on startup
    Set-Content $Path {
            If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
            $Arguments = "& '" + $MyInvocation.mycommand.definition + "'"
            Start-Process powershell -Verb runAs -ArgumentList $arguments
            Break
        } 
        Remove-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"               
        Start-Process -Filepath "C:\Exchange Resources\UcmaRuntimeSetup.exe" -ArgumentList "/passive /norestart" -wait
        Start-Sleep -s 15
        Start-Process -FilePath "C:\Exchange Resources\NDP452-KB2901907-x86-x64-AllOS-ENU.exe" -ArgumentList "/passive /norestart" -wait
        Start-Sleep -s 15                
        Start-Process -Filepath "C:\Exchange Resources\filterpack2010sp2-x64-fullfile-en-us.exe" -ArgumentList "/passive /norestart" -wait
        Start-Sleep -s 15
        $Path = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
        $TestPath = Test-Path "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
        if( $TestPath -eq $false ) {
            New-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1" -type file
        }
        else {
            Write-Host "Script already Exists in $env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\" `n
            Write-Host "Resuming Script..." `n
        }
 # Sets Content for InstallExchange.ps1 script upon 2nd reboot to run on startup
        Set-Content $Path {
            If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
                $Arguments = "& '" + $MyInvocation.mycommand.definition + "'"
                Start-Process powershell -Verb runAs -ArgumentList $arguments
                Break
            }        
            $Organization = Get-Content -Path "C:\Exchange Resources\OrganizationName.txt" 
            Write-Host "Server is now ready to install Exchange Server 2016..."
            Function Enable-OpenFileSecurityWarning {
            Write-Output "Enabling File Security Warning dialog..."
            Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations" -Name "LowRiskFileTypes" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments" -Name "SaveZoneInformation" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations" -Name "LowRiskFileTypes" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments" -Name "SaveZoneInformation" -ErrorAction SilentlyContinue
            Write-Output "Finished!"
            }
            Function Enable-IEESC {
            Write-Verbose "Enabling IE Enhanced Security Configuration..."
            $AdminKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"
            $UserKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}"
            Set-ItemProperty -Path $AdminKey -Name "IsInstalled" -Value 1
            Set-ItemProperty -Path $UserKey -Name "IsInstalled" -Value 1
            Stop-Process -Name Explorer
            Write-Output "Finished!"
            }
            Function Enable-UAC {
            Write-Output "Enabling User Account Control..."
            New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 1 -ErrorAction SilentlyContinue| out-null
            Write-Output "Finished!"
            }
            C:\Exchange_2016_Setup\setup.exe /PrepareSchema /IAcceptExchangeServerLicenseTerms
            C:\Exchange_2016_Setup\setup.exe /PrepareAD /OrganizationName:"$Organization" /IAcceptExchangeServerLicenseTerms
            C:\Exchange_2016_Setup\setup.exe /PrepareAllDomains /IAcceptExchangeServerLicenseTerms                       
            C:\Exchange_2016_Setup\Setup.exe /Mode:Install /Role:Mailbox /OrganizationName:"$Organization" /TargetDir:"C:\Exchange Server 2016" /IAcceptExchangeServerLicenseTerms
 # Changes file association for .ps1 to default (notepad.exe)
            Set-ItemProperty -Path HKLM:\SOFTWARE\Classes\Microsoft.PowerShellScript.1\Shell\Open\Command -Name "(Default)" -Value '"C:\Windows\System32\notepad.exe" "%1"' -Type String -Force
 # Changes QuickEdit to default registry setting
            Set-ItemProperty -Path HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe -Name "QuickEdit" -Value '1' -Type DWord -Force
 # Changes DoNotOpenServerManagerAtLogon to default registry setting
            New-ItemProperty -Path HKCU:\Software\Microsoft\ServerManager -Name "DoNotOpenServerManagerAtLogon" -Value "0" -Type DWord -Force
 # Removes InstallExchange.ps1 script from startup folder
            Remove-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
 # Removes leftover text file for Organization Name
            Remove-Item "C:\Exchange Resources\OrganizationName.txt"
 # Re-enables UAC
            Enable-UAC
 # Re-enables IEESC
            Enable-IEESC
 # Re-enables OpenFileSecurityWarning
            Enable-OpenFileSecurityWarning
 # Changes AutoAdminLogon to default  registry setting
            $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"  
            Set-ItemProperty $RegPath "AutoAdminLogon" -Value "0" -type String
            Pause
            Restart-Computer
          }
    Restart-Computer
    }
Restart-Computer


}
4{ 
$a = (Get-Host).UI.RawUI
$a.BackgroundColor = "black"
$a.ForegroundColor = "green"

}
5{
}
6{
}
}
}
2 {
while ( $xMenu2 -lt 1 -or $xMenu2 -gt 5 ){
CLS
# Present the Menu Options
Write-Host “`n`tBrads Ultimate Automation Script`n” -Fore Magenta
Write-Host “`t`t Please select the Distribution Group task you require`n” -Fore Green
Write-Host “`t`t`t1. Create a New Distribution Group” -Fore Green
Write-Host “`t`t`t2. Add Member to an Existing Distribution Group” -Fore Green
Write-Host “`t`t`t3. Remove Member from an Existing Distribution Group” -Fore Green
Write-Host “`t`t`t4. Go to Main Menu`n” -Fore Green
[int]$xMenu2 = Read-Host “`tEnter Menu Option Number”
}
if( $xMenu1 -lt 1 -or $xMenu1 -gt 4 ){

Write-Host “`tPlease select one of the options available.`n” -Fore Red;start-Sleep -Seconds 1
}
Switch ($xMenu2){
1{ 

Write-Host “`n`t##############################################################" -Fore Green
Write-Host "`t#                INCLUDE DOMAIN IN USERNAME                  #" -Fore Green
Write-Host "`t##############################################################" -Fore Green

$a = (Get-Host).UI.RawUI
$a.BackgroundColor = "black"
$a.ForegroundColor = "magenta"

}
2{ 

Write-Host “`n`t##############################################################" -Fore Green
Write-Host "`t#                INCLUDE DOMAIN IN USERNAME                  #" -Fore Green
Write-Host "`t##############################################################" -Fore Green

$a = (Get-Host).UI.RawUI
$a.BackgroundColor = "black"
$a.ForegroundColor = "magenta"


}

3{ 

Write-Host “`n`t##############################################################" -Fore Green
Write-Host "`t#                INCLUDE DOMAIN IN USERNAME                  #" -Fore Green
Write-Host "`t##############################################################" -Fore Green

$a = (Get-Host).UI.RawUI
$a.BackgroundColor = "black"
$a.ForegroundColor = "magenta"


}
4{ Write-Host “`n`tYou Selected Option 4 – Go to Main Menu`n” -Fore Yellow; break}
}
}
3 {
while ( $xMenu2 -lt 1 -or $xMenu2 -gt 4 ){
CLS
# Present the Menu Options
Write-Host “`n`tBrads Ultimate Automation Script`n” -Fore Magenta
Write-Host “`t`tPlease select a mailbox administration task`n” -Fore Green
Write-Host “`t`t`t1. Create a New Mailbox” -Fore Green
Write-Host “`t`t`t2. Delete a Mailbox” -Fore Green
Write-Host “`t`t`t3. Create a New Mail Contact” -Fore Green
Write-Host “`t`t`t4. Go to Main Menu`n” -Fore Green
[int]$xMenu2 = Read-Host “`tEnter Menu Option Number”
if( $xMenu1 -lt 1 -or $xMenu1 -gt 4 ){
Write-Host “`tPlease select one of the options available.`n” -Fore Red;start-Sleep -Seconds 1
}
}
Switch ($xMenu2){
1{ 
Write-Host “`n`t##############################################################" -Fore Green
Write-Host "`t#                INCLUDE DOMAIN IN USERNAME                  #" -Fore Green
Write-Host "`t##############################################################" -Fore Green

$a = (Get-Host).UI.RawUI
$a.BackgroundColor = "black"
$a.ForegroundColor = "magenta"

}
2{ 

Write-Host “`n`t##############################################################" -Fore Green
Write-Host "`t#                INCLUDE DOMAIN IN USERNAME                  #" -Fore Green
Write-Host "`t##############################################################" -Fore Green

$a = (Get-Host).UI.RawUI
$a.BackgroundColor = "black"
$a.ForegroundColor = "magenta"

 }
3{ 

Write-Host “`n`t##############################################################" -Fore Green
Write-Host "`t#                INCLUDE DOMAIN IN USERNAME                  #" -Fore Green
Write-Host "`t##############################################################" -Fore Green

$a = (Get-Host).UI.RawUI
$a.BackgroundColor = "black"
$a.ForegroundColor = "magenta"

 }
4{ Write-Host “`n`tYou Selected Option 4 – Go to Main Menu`n” -Fore Yellow; break}
}
}
4 {
while ( $xMenu2 -lt 1 -or $xMenu2 -gt 12 ){
CLS
# Present the Menu Options
Write-Host “`n`tBrads Ultimate Automation Script`n” -Fore Magenta
Write-Host “`t`tPlease select the Software Administration Task`n” -Fore Green
Write-Host “`t`tSilent Installs`n” -Fore Green
Write-Host “`t`t`t1. Install Microsoft Office 2010 32-bit” -Fore Green
Write-Host “`t`t`t2. Install Microsoft Office 2010 64-bit” -Fore Green
Write-Host “`t`t`t3. Install Microsoft Office 2010 Language Pack” -Fore Green
Write-Host “`t`t`t4. Install Microsoft Project 2010 32-bit” -Fore Green
Write-Host “`t`t`t5. Install Microsoft Visio 2010 32-bit” -Fore Green
Write-Host “`t`t`t6. Microsoft Visual Studio 2010 Professional Full`n” -Fore Green
Write-Host “`t`tRegular Installs`n” -Fore Green
Write-Host “`t`t`t7. Install Microsoft Office 2010 32-bit” -Fore Green
Write-Host “`t`t`t8. Install Microsoft Office 2010 64-bit” -Fore Green
Write-Host “`t`t`t9. Install Microsoft Office 2010 Language Pack” -Fore Green
Write-Host “`t`t`t10. Install Microsoft Project 2010 32-bit” -Fore Green
Write-Host “`t`t`t11. Install Microsoft Visio 2010 32-bit`n” -Fore Green
Write-Host “`t`t`t12. Go to Main Menu`n” -Fore Green
[int]$xMenu2 = Read-Host “`tEnter Menu Option Number”
if( $xMenu1 -lt 1 -or $xMenu1 -gt 12 ){
Write-Host “`tPlease select one of the options available.`n” -Fore Red;start-Sleep -Seconds 1
}
}
Switch ($xMenu2){
1{ & '\\filer\Software\IT\Microsoft\MS Office\MS Office 2010 Standard SP1\setup.exe' /config "\\filer\Software\IT\Microsoft\MS Office\MS Office 2010 Standard SP1\Standard.WW\configv2.xml"}
2{ & '\\filer\Software\IT\Microsoft\MS Office\MS Office 2010 Standard x64\setup.exe' /config "\\filer\Software\IT\Microsoft\MS Office\MS Office 2010 Standard x64\Standard.WW\configv2.xml"}
3{ & '\\filer\software\Apps\Office2010 Language Pack\Extracted\setup.exe' /config "\\filer\Software\Apps\Office2010 Language Pack\Extracted\ProofKit.WW\configv2.xml"}
4{ & '\\filer\Software\IT\Microsoft\MS Project\MS Project 2010 Standard w SP1\setup.exe' /config "\\filer\Software\IT\Microsoft\MS Project\MS Project 2010 Standard w SP1\PrjStd.WW\configv2.xml"}
5{ & '\\filer\Software\IT\Microsoft\MS Visio\MS Visio 2010 w SP1\x86\setup.exe' /config "\\filer\Software\IT\Microsoft\MS Visio\MS Visio 2010 w SP1\x86\Visio.WW\configv2.xml"}
6{ & '\\filer\software\IT\Microsoft\MS_Visual_Studio\VS2010Professional\extracted\Setup\setup.exe' /unattendfile "\\filer\Documents\IT\_Scripts\_Script Resource Data\VS2010ProfessionalDeployment.ini"}
7{ & '\\filer\Software\IT\Microsoft\MS Office\MS Office 2010 Standard SP1\setup.exe' /config "\\filer\Software\IT\Microsoft\MS Office\MS Office 2010 Standard SP1\Standard.WW\config.xml"}
8{ & '\\filer\Software\IT\Microsoft\MS Office\MS Office 2010 Standard x64\setup.exe' /config "\\filer\Software\IT\Microsoft\MS Office\MS Office 2010 Standard x64\Standard.WW\config.xml"}
9{ & '\\filer\software\Apps\Office2010 Language Pack\Extracted\setup.exe' /config "\\filer\Software\Apps\Office2010 Language Pack\Extracted\ProofKit.WW\config.xml"}
10{ & '\\filer\Software\IT\Microsoft\MS Project\MS Project 2010 Standard w SP1\setup.exe' /config "\\filer\Software\IT\Microsoft\MS Project\MS Project 2010 Standard w SP1\PrjStd.WW\config.xml"}
11{ & '\\filer\Software\IT\Microsoft\MS Visio\MS Visio 2010 w SP1\x86\setup.exe' /config "\\filer\Software\IT\Microsoft\MS Visio\MS Visio 2010 w SP1\x86\Visio.WW\config.xml"}
12{ Write-Host “`n`tYou Selected Option 10 – Go to Main Menu`n” -Fore Yellow; break}
}
}
default { $global:xExitSession=$true;break }
}
}
LoadMenuSystem
If ($xExitSession){
Exit-PSSession    #… User quit & Exit
} Else {
"C:\Users\Brad Stevens\AppData\Roaming\Microsoft\Windows\Network Shortcuts\OneDrive Powershell Scripts\Menu.ps1"    #… Loop the function
}