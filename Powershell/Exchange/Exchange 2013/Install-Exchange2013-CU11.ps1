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
        Updated the script to include CU11 instead of CU10
        Updated script to include 2010 Filter Pack SP2
        Added check for already installed prerequisite packages
    1.3 Created Welcome Window
    1.4 Added folder browsing function
    1.5 Added switch function
    1.6 Added previous Exchange application download location inquiry
    1.7 Added Exchange download location inquiry
    1.8 Added Exchange install directory inquiry
    1.9 Adjusted package buffer byte rate
    2.0 Debugged console window output
  


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
        $Package_Buffer = new-object byte[] 100000KB
        $Package_Count = $Package_ResponseStream.Read($Package_Buffer,0,$Package_Buffer.length)
        $Package_DownloadedBytes = $Package_Count
        while ($Package_Count -gt 0) {
            $Package_Stream.Write($Package_Buffer, 0, $Package_Count)
            $Package_Count = $Package_ResponseStream.Read($Package_Buffer,0,$Package_Buffer.length)
            $Package_DownloadedBytes = $Package_DownloadedBytes + $Package_Count
            Write-Progress -activity "Downloading file '$($Package_Url.split('/') | select -Last 1)'" -status "Downloaded ($([System.Math]::Floor($Package_DownloadedBytes/1024))K of $($Package_Total)K): " -PercentComplete ((([System.Math]::Floor($Package_DownloadedBytes/1024)) / $Package_Total)  * 100)
         }
        Write-Progress -activity "Finished downloading file '$($Package_Url.split('/') | select -Last 1)'"
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
    function Read-FolderBrowserDialog([string]$Message, [string]$InitialDirectory, [switch]$NoNewFolderButton){
    $browseForFolderOptions = 0
    if ($NoNewFolderButton) { $browseForFolderOptions += 512 }
 
    $app = New-Object -ComObject Shell.Application
    $folder = $app.BrowseForFolder(0, $Message, $browseForFolderOptions, $InitialDirectory)
    if ($folder) { $selectedDirectory = $folder.Self.Path } else { $selectedDirectory = '' }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) > $null
    return $selectedDirectory
    }
    function Read-OpenFileDialog([string]$WindowTitle, [string]$InitialDirectory, [string]$Filter = "All files (*.*)|*.*", [switch]$AllowMultiSelect){  
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = $WindowTitle
    if (![string]::IsNullOrWhiteSpace($InitialDirectory)) { $openFileDialog.InitialDirectory = $InitialDirectory }
    $openFileDialog.Filter = $Filter
    if ($AllowMultiSelect) { $openFileDialog.MultiSelect = $true }
    $openFileDialog.ShowHelp = $true    # Without this line the ShowDialog() function may hang depending on system configuration and running from console vs. ISE.
    $openFileDialog.ShowDialog() > $null
    if ($AllowMultiSelect) { return $openFileDialog.Filenames } else { return $openFileDialog.Filename }
}


   #####################################################################################
   #                                                                                   #
   #                                 Main Script                                       #
   #                                                                                   #
   #####################################################################################
    $a = (Get-Host).UI.RawUI
    $a.BackgroundColor = "black"
    $a.ForegroundColor = "White"
    $c = $a.BufferSize
    $c.Width = 125
    $c.Height = 35
    $a.BufferSize = $c
    $b = $a.WindowSize
    $b.Width = 125
    $b.Height = 35
    $a.WindowSize = $b

    CLS

    If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
    $Arguments = "& '" + $MyInvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
    }
    Disable-UAC
    Disable-OpenFileSecurityWarning
    Disable-IEESC
    Clear-Host
    Write-Host "   __        __   _                            _          _   _            _____          _                            "
    Write-Host "   \ \      / /__| | ___ ___  _ __ ___   ___  | |_ ___   | |_| |__   ___  | ____|_  _____| |__   __ _ _ __   __ _  ___ "
    Write-Host "    \ \ /\ / / _ \ |/ __/ _ \| '_ ` _  \ / _ \ | __/ _ \  | __| '_ \ / _ \ |  _| \ \/ / __| '_ \ / _`  | '_ \ / _`  |/ _ \"
    Write-Host "     \ V  V /  __/ | (_| (_) | | | | | |  __/ | || (_) | | |_| | | |  __/ | |___ >  < (__| | | | (_| | | | | (_| |  __/"
    Write-Host "      \_/\_/ \___|_|\___\___/|_| |_| |_|\___|  \__\___/   \__|_| |_|\___| |_____/_/\_\___|_| |_|\__,_|_| |_|\__, |\___|"
    Write-Host "                                                                                                            |___/      "
    Write-Host "       _         _                        _   _               ____            _       _   _ "
    Write-Host "      / \  _   _| |_ ___  _ __ ___   __ _| |_(_) ___  _ __   / ___|  ___ _ __(_)_ __ | |_| |"
    Write-Host "     / _ \| | | | __/ _ \| '_ ` _  \ / _`  | __| |/ _ \| '_ \  \___ \ / __| '__| | '_ \| __| |"
    Write-Host "    / ___ \ |_| | || (_) | | | | | | (_| | |_| | (_) | | | |  ___) | (__| |  | | |_) | |_|_|"
    Write-Host "   /_/   \_\__,_|\__\___/|_| |_| |_|\__,_|\__|_|\___/|_| |_| |____/ \___|_|  |_| .__/ \__(_)"
    Write-Host "                                                                               |_|          "
    Write-Host `n
    Write-Host `n
    Write-Host `n
    Write-Host `t "Created By: Brad C. Stevens"
    Write-Host `t "Associate Consulting Engineer, CDW"
    Write-Host `n
    Write-Host `n
    Write-Host `n
    Write-Host `t "Press any key to begin..."
    $HOST.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
    $HOST.UI.RawUI.Flushinputbuffer()
    Clear-Host
    Write-Host "Checking server domain membership....." `n
    If((gwmi win32_computersystem).partofdomain -eq $true) {
    Write-Host "Server is Domain Joined! Resuming Script...."
    Write-Host `n
    Write-Host "Prompt for Domain Credentials..."
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
    Clear-Host
# Loads System.Reflection.Assembly to prompt user for Exchange Organization Name
    $OrganizationName= Read-Host "Please enter an Exchange Organization name"
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
    $Resources = Test-Path "C:\Exchange Resources"
    if($Resources -eq $false) {
        New-Item "C:\Exchange Resources" -type directory
    }

    Clear-Host
    

    $title = "Exchange Executable Location"
    $message = "Have you already downloaded Exchange CU11?"
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
   :OuterLoop Do {
   $result = $host.ui.PromptForChoice($title, $message, $options, 0) 
    switch ($result)
        {
            0 {
                $DownloadDirectory = Read-OpenFileDialog -WindowTitle "Exchange Directory" -InitialDirectory 'C:\'
                if (![string]::IsNullOrEmpty($DownloadDirectory)) { 
                Write-Host `n
                Write-Host "You selected the file: $DownloadDirectory"
                Write-Host `n
                $ExchangeInstallLocation = Read-FolderBrowserDialog -Message "Please select the location where you would like to have Exchange installed"
                if (![string]::IsNullOrEmpty($ExchangeInstallLocation)) { Write-Host "You selected the directory: $ExchangeInstallLocation" -ForegroundColor "Green" }
                else {                
                }}
                else { CLS
                       Write-Host `n
                       Write-Host "You did not select a location. Please select the Exchange Executable file" -foreground "Red"
                       $DownloadDirectory = $null    
                 }   
            }
            1 {
                    CLS
                    Write-Host "Please choose a directory for Exchange Resources"
                    $Directory = Read-FolderBrowserDialog -Message "Please select a directory"
                    if (![string]::IsNullOrEmpty($Directory)) { Write-Host "You selected the directory: '$Directory'" 
                        $DownloadDirectory = $Directory + "\Exchange2013-x64-cu11.exe"
                        Write-Host "Downloading Installation Packages..." `n 
                        Package-Download "https://download.microsoft.com/download/A/A/B/AAB18934-BC8F-429D-8912-6A98CBC96B07/Exchange2013-x64-cu11.exe" "$DownloadDirectory" 
                        Write-Host "Finished downloading Exchange2013-x64-cu11.exe"
                        $ExchangeInstallLocation = Read-FolderBrowserDialog -Message "Please select the location where you would like to have Exchange installed"
                        if (![string]::IsNullOrEmpty($ExchangeInstallLocation)) { Write-Host "You selected the directory: $ExchangeInstallLocation" -ForegroundColor "Green" }
                        else {                
                        }
                        
                        }
                    else { 
                        CLS
                        Write-Host `n 
                        Write-Host "You did not select a directory." -foreground "Red"
                        $Directory = $Null
                    }
                }


            }
        }
        While ($DownloadDirectory -eq $Null )
    CLS
    $OrganizationFilePath = Test-Path "C:\Exchange Resources\OrganizationName.txt"
    if($OrganizationFilePath -eq $false) {
        New-Item "C:\Exchange Resources\OrganizationName.txt" -ItemType File
        Set-Content -Path "C:\Exchange Resources\OrganizationName.txt" -Value "$OrganizationName"
    }
    $InstallPath = Test-Path "C:\Exchange Resources\InstallPath.txt"
    if($InstallPath -eq $false) {
        New-Item "C:\Exchange Resources\InstallPath.txt" -ItemType File
        Set-Content -Path "C:\Exchange Resources\InstallPath.txt" -Value "$ExchangeInstallLocation"
    }
    CLS
    Write-Host `n
    Write-Host `n
    Write-Host `n
    Write-Host `n
 
# Download request for Exchange 2013 CU9
    Write-Host "Downloading Filter Pack SP1..."
# Download request for 2010 Filter Pack SP2``
    $FilterPackSp2 = Test-Path "C:\Exchange Resources\filterpack2010sp2-x64-fullfile-en-us.exe"
    If($FilterPackSp2 -eq $false) {
    Package-Download "http://download.microsoft.com/download/D/C/A/DCA32A51-6954-4814-8838-422BD3F508F8/filterpacksp2010-kb2687447-fullfile-x64-en-us.exe" "C:\Exchange Resources\filterpack2010sp2-x64-fullfile-en-us.exe"
    Write-Host "Done!" -ForegroundColor "Green"
    Write-Host "Downloading UCMA Runtime..."
    }
# Download request for UcmaRuntimeSetup
    $UCMARuntimeSetup = Test-Path "C:\Exchange Resources\UcmaRuntimeSetup.exe"
    If($UCMARuntimeSetup -eq $false) {
    Package-Download "http://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe" "C:\Exchange Resources\UcmaRuntimeSetup.exe"
    Write-Host "Done!" -ForegroundColor "Green"
    Write-Host "Downloading .NET Framework 4.5.2..."
    }
# Download request for .NET Framework 4.5.2
    $NDP452 = Test-Path "C:\Exchange Resources\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
    If($NDP452 -eq $false) {
    Package-Download "http://download.microsoft.com/download/E/2/1/E21644B5-2DF2-47C2-91BD-63C560427900/NDP452-KB2901907-x86-x64-AllOS-ENU.exe" "C:\Exchange Resources\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
    Write-Host "Done!" -ForegroundColor "Green"
    Write-Host "Installing .NET Framework 4.5.2..."
    }
# Installation of Windows Feature Prerequisites
    Start-Process -FilePath "C:\Exchange Resources\NDP452-KB2901907-x86-x64-AllOS-ENU.exe" -ArgumentList "/passive /norestart" -wait
    Write-Host "Done!" -ForegroundColor "Green"
    Write-Host "Installing Windows Features..."
    Install-WindowsFeature RSAT-ADDS
    Install-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
    Write-Host "Done!" -ForegroundColor "Green"
# Specifys Extraction Folder for Exchange Setup
    $Targetfolder="C:\Exchange_2013_CU11_Setup"
# Checks, then Extracts Exchange to Extraction Folder
    $Extract = Test-Path "C:\Exchange_2013_CU11_Setup"
    if ($Extract -eq $false) {
        Write-Host "Extracting Exchange..."
        Start-Process -Filepath "C:\Exchange Resources\Exchange2013-x64-cu11.exe" -ArgumentList "/extract:$Targetfolder /u" -Wait
        Write-Host "Done!" -ForegroundColor "Green"         
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
            $ExchangeInstallLocation = Get-Content -Path "C:\Exchange Resources\InstallPath.txt" 
            Write-Host "Server is now ready to install Exchange Server 2013 CU11..."
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
            C:\Exchange_2013_CU11_Setup\setup.exe /PrepareSchema /IAcceptExchangeServerLicenseTerms
            C:\Exchange_2013_CU11_Setup\setup.exe /PrepareAD /OrganizationName:"$Organization" /IAcceptExchangeServerLicenseTerms                     
            C:\Exchange_2013_CU11_Setup\Setup.exe /mode:Install /role:ClientAccess,Mailbox /TargetDir:"$ExchangeInstallLocation" /OrganizationName:"$Organization" /IAcceptExchangeServerLicenseTerms          
 # Changes file association for .ps1 to default (notepad.exe)
            Set-ItemProperty -Path HKLM:\SOFTWARE\Classes\Microsoft.PowerShellScript.1\Shell\Open\Command -Name "(Default)" -Value '"C:\Windows\System32\notepad.exe" "%1"' -Type String -Force
 # Changes QuickEdit to default registry setting
            Set-ItemProperty -Path HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe -Name "QuickEdit" -Value '1' -Type DWord -Force
 # Changes DoNotOpenServerManagerAtLogon to default registry setting
            New-ItemProperty -Path HKCU:\Software\Microsoft\ServerManager -Name "DoNotOpenServerManagerAtLogon" -Value "0" -Type DWord -Force
 # Removes InstallExchange.ps1 script from startup folder
            Remove-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
 # Removes leftover text file for Organization Name
            Remove-Item "C:\Exchange Resources" -Recurse -Force
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