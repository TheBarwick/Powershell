<#
    .TITLE
    Brad's Ultimate Exchange Server 2016 RTM Automation Script
    
    .SYNOPSIS
    Installs Exchange Server 2016 RTM Client Access and Mailbox Roles and all required prerequisites.
    
    Adapted from:
    - http://blogs.msdn.com/b/jasonn/archive/2013/06/11/8594493.aspx;
    - http://eightwone.com/2013/02/18/exchange-2013-unattended-installation-script
    - https://technet.microsoft.com/en-us/library/ff730939.aspx



    Brad C. Stevens
    brad@bradcstevens.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 2.0, Janurary 13th, 2016

    .DESCRIPTION
    This script will install Exchange 2016 RTM prerequisites, creates the Exchange 
    organization (prepares Active Directory) and installs Exchange Server. 


    .LINK
    - http://bradcstevens.com
    - http://helpdesksage.blogspot.com

    .GITHUB
    https://github.com/bradcstevens

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
        Updated the script to include RTM 2016 instead of 2013
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

    Function Set-Console {
    # Forces PowerShell Console to run with Administrative Rights
        If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")){   
        $Arguments = "& '" + $MyInvocation.mycommand.definition + "'"
        Start-Process powershell -Verb runAs -ArgumentList $arguments
        Break
        }
    }
    
    Function Display-WelcomeScreen {
    Clear-Host
    Write-Host `n `n `n `t "Welcome to the Unattended Exchange Installation Script!" 
    Write-Host `n `n `n `t "Version:" -ForegroundColor "Yellow" -NoNewLine
    Write-Host " 2.0"
    Write-Host `t "Cumulative Update:" -ForegroundColor "Yellow" -NoNewLine 
    Write-Host " 12"
    Write-Host `n `t "Created By:" -ForegroundColor "Yellow" -NoNewline 
    Write-Host " Brad C. Stevens"
    Write-Host `t "Title:" -ForegroundColor "Yellow" -NoNewLine 
    Write-Host " Associate Consulting Engineer, CDW"
    Write-Host `n `t "Requirements:" -ForegroundColor "Yellow" -NoNewline
    Write-Host    "  - PowerShell v3;" 
    Write-Host `t "               - Windows Server 2008 R2/2012/2012 R2;" 
    Write-Host `t "               - Domain-joined system;"
    Write-Host `t "               - System is in same site as Schema Master"  
    Write-Host `t "               - Domain account with Domain Admin, Enterprise Admin, and Schema Admin membership" 
    Write-Host `n `t "Press any key to begin setup..." -ForegroundColor "Yellow"
    # Pauses script until key press
    $HOST.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL $HOST.UI.RawUI.Flushinputbuffer()
    Clear-Host 
    }
     Function Display-ReadyScreen {
    $OrganizationName = Get-Content -Path "C:\Exchange_Resources\OrganizationName.txt"
    $ExchangeInstallLocation = Get-Content -Path "C:\Exchange_Resources\InstallPath.txt"
    $DomainName = Get-Content -Path "C:\Exchange_Resources\DomainName.txt"
    $DomainUserName = Get-Content -Path "C:\Exchange_Resources\DomainUserName.txt"
    Clear-Host
    Write-Host `n `n `n `t "Script is ready to start!" 
    Write-Host `n `n `n `t "Your Organization Name will be: " -ForegroundColor "Yellow" -NoNewLine
    Write-Host "$OrganizationName"
    Write-Host `n `t "Exchange 2016 RTM will install in the directory:" -ForegroundColor "Yellow" -NoNewLine 
    Write-Host " $ExchangeInstallLocation"
    Write-Host `n `t "Domain:" -ForegroundColor "Yellow" -NoNewline 
    Write-Host " $DomainName"
    Write-Host `n `t "Username:" -ForegroundColor "Yellow" -NoNewLine 
    Write-Host " $DomainUserName"
    Write-Host `n `n
    Write-Host `n `t "Press any key to begin unattended install..."
    # Pauses script until key press
    $HOST.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL $HOST.UI.RawUI.Flushinputbuffer()
    Clear-Host
    }

    Function Check-DomainMembership {
        Write-Host "Checking server domain membership....." `n
        # Checks if system is domain joined
        If((gwmi win32_computersystem).partofdomain -eq $true) {
        Write-Host "Server is domain joined!" -ForegroundColor "yellow" `n
        Write-Host "Prompting for Domain Credentials..." `n
        }   
        Else {
        Write-Host -fore red "Domain Membership Required. Join this server to a domain then run the script."
        Pause
        Exit
        }
    }

    Function Set-AutomaticLogon {
    $Credentials = $host.ui.PromptForCredential("Domain Administrator Credentials", "Please enter your domain user name and password in the following format: domain\username.", "", "NetBiosUserName")
    $Domain = $Credentials.GetNetworkCredential().Domain 
    $PlainUsername = $Credentials.GetNetworkCredential().UserName
    $PlainPassword = $Credentials.GetNetworkCredential().Password
    $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"  
    Set-ItemProperty $RegPath "AutoAdminLogon" -Value "1" -type String  
    Set-ItemProperty $RegPath "DefaultUsername" -Value "$Domain\$PlainUsername" -type String  
    Set-ItemProperty $RegPath "DefaultPassword" -Value "$PlainPassword" -type String
    Set-ItemProperty -Path HKLM:\SOFTWARE\Classes\Microsoft.PowerShellScript.1\Shell\Open\Command -Name "(Default)" -Value '"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" "-file" "%1"' -Type String -Force
    Set-ItemProperty -Path HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe -Name "QuickEdit" -Value '0' -Type DWord -Force
    New-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage -Name "OpenAtLogon" -Value '0' -Type DWord -Force
    New-ItemProperty -Path HKCU:\Software\Microsoft\ServerManager -Name "DoNotOpenServerManagerAtLogon" -Value "1" -Type DWord -Force
    New-Item "C:\Exchange_Resources\DomainUserName.txt" -ItemType File
    Set-Content -Path "C:\Exchange_Resources\DomainUserName.txt" -Value "$PlainUsername"
    New-Item "C:\Exchange_Resources\DomainName.txt" -ItemType File
    Set-Content -Path "C:\Exchange_Resources\DomainName.txt" -Value "$Domain"
    Clear-Host
    Write-Host "Credentials have been set for $PlainUserName's $Domain domain account!"-ForegroundColor "Yellow" `n          
    }

    Function Package-Download($Package_Url, $Package_File) {
        $Package_Stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Package_File, Create
        $Package_URI = New-Object "System.Uri" "$Package_Url"
        $Package_Request = [System.Net.HttpWebRequest]::Create($Package_URI)
        # 30 Second Timeout
        $Package_Request.set_Timeout(30000)
        $Package_Response = $Package_Request.GetResponse()
        $Package_Total = [System.Math]::Floor($Package_Response.get_ContentLength()/1024)
        $Package_ResponseStream = $Package_Response.GetResponseStream()
        $Package_Buffer = new-object byte[] 1MB
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
    Function Disable-FireWall(){
        $status = netsh advfirewall show allprofiles state
        If ($status | Select-String "ON") {
            $enabled = $true
        }
        Else {
        $enabled = $false
        }
        If ($enabled -eq $true) {
            netsh advfirewall set allprofiles state off
            Write-Host "Firewall is now disabled" -ForegroundColor yellow
            return
        }
        If ($enabled -eq $false) {
            Write-Host "Firewall is already disabled" -ForegroundColor yellow
        }
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
    Function Read-FolderBrowserDialog([string]$Message, [string]$InitialDirectory, [switch]$NoNewFolderButton){
    $browseForFolderOptions = 0
        If($NoNewFolderButton){ 
            $browseForFolderOptions += 512 
        }
    $app = New-Object -ComObject Shell.Application
    $folder = $app.BrowseForFolder(0, $Message, $browseForFolderOptions, $InitialDirectory)
        If ($folder) { 
        $selectedDirectory = $folder.Self.Path } else { $selectedDirectory = '' }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) > $null
        return $selectedDirectory
        }
    
    Function Read-OpenFileDialog([string]$WindowTitle, [string]$InitialDirectory, [string]$Filter = "All files (*.*)|*.*", [switch]$AllowMultiSelect){  
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = $WindowTitle
        If(![string]::IsNullOrWhiteSpace($InitialDirectory)){ 
            $openFileDialog.InitialDirectory = $InitialDirectory
        }
    $openFileDialog.Filter = $Filter
        If($AllowMultiSelect){ 
        $openFileDialog.MultiSelect = $true 
        }
    $openFileDialog.ShowHelp = $true    # Without this line the ShowDialog() function may hang depending on system configuration and running from console vs. ISE.
    $openFileDialog.ShowDialog() > $null
        If($AllowMultiSelect){ 
        return $openFileDialog.Filenames
        } 
        Else{ 
        return $openFileDialog.Filename 
        }
    }
    Function Get-InstalledApps {
        if ([IntPtr]::Size -eq 4) {
            $regpath = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
        }
        else {
            $regpath = @(
                'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
                'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
            )
        }
        Get-ItemProperty $regpath | .{process{if($_.DisplayName -and $_.UninstallString) { $_ } }} | Select DisplayName, Publisher, InstallDate, DisplayVersion, UninstallString |Sort DisplayName
    }

    Function Create-ResourcesDirectory {
        $Resources = Test-Path "C:\Exchange_Resources"
        if($Resources) {
        }
        else {
            New-Item "C:\Exchange_Resources" -type directory
        }
    }

    Function Save-OrganizationName {
        Clear-Host  
        $OrganizationName = Read-Host "Please enter a name for your Exchange Organization"
        Write-Host `n
    # RegEx to remove spaces in Organization Name 
        $OrganizationName = $OrganizationName -replace '\s+', '' 
        Write-Host "Your Organization Name will be: '$OrganizationName'" -ForegroundColor "Yellow" `n
        $OrganizationFilePath = Test-Path "C:\Exchange_Resources\OrganizationName.txt"
        if($OrganizationFilePath) {
        }
        else {
            New-Item "C:\Exchange_Resources\OrganizationName.txt" -ItemType File
            Set-Content -Path "C:\Exchange_Resources\OrganizationName.txt" -Value "$OrganizationName"
        }
    }

    Function Install-FilterPack2 {
        Write-Host `n `n `n `n "Downloading Filter Pack SP2..." `n
        # Download request for 2010 Filter Pack SP2
        $FilterPackSp2InstallCheck = '*Filter Pack SP2*'
        $FilterPackSp2Installed = Get-InstalledApps | where {$_.DisplayName -like $FilterPackSp2InstallCheck}
        Clear-Host
        If ($FilterPackSp2Installed -eq $null) {
            $FilterPackSp2 = Test-Path "C:\Exchange_Resources\filterpack2010sp2-x64-fullfile-en-us.exe"
            If($FilterPackSp2 -eq $false) {
                Package-Download "http://download.microsoft.com/download/D/C/A/DCA32A51-6954-4814-8838-422BD3F508F8/filterpacksp2010-kb2687447-fullfile-x64-en-us.exe" "C:\Exchange_Resources\filterpack2010sp2-x64-fullfile-en-us.exe"
                Write-Host "Done!" -ForegroundColor "Yellow" `n
                Write-Host "Scanning for UCMA Runtime..." `n
            }
            ElseIf($FilterPackSp2 -eq $true) {
                $FilterPackSp2 = "C:\Exchange_Resources\filterpack2010sp2-x64-fullfile-en-us.exe"
                Remove-Item $FilterPackSp2 -Force
                Clear-Host
                Write-Host "Downloading Filter Pack SP2..." `n
                Package-Download "http://download.microsoft.com/download/D/C/A/DCA32A51-6954-4814-8838-422BD3F508F8/filterpacksp2010-kb2687447-fullfile-x64-en-us.exe" "C:\Exchange_Resources\filterpack2010sp2-x64-fullfile-en-us.exe"
                Write-Host "Done!" -ForegroundColor "Yellow" `n
                Write-Host "Scanning for UCMA Runtime..." `n
                Write-Host "Scanning for UCMA Runtime..." -ForegroundColor "Yellow"
            }
        }
    }

    Function Install-UCMA {
        $UCMAInstallCheck = '*UCMA Runtime*'
        $UCMAInstalled = Get-InstalledApps | where {$_.DisplayName -like $UCMAInstallCheck}
        Clear-Host
        If ($UCMAInstalled -eq $null) {
            Write-Host "Downloading UCMA Runtime Setup..." `n
            $UCMARuntimeSetup = Test-Path "C:\Exchange_Resources\UcmaRuntimeSetup.exe"
            If($UCMARuntimeSetup -eq $false) {
                Package-Download "http://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe" "C:\Exchange_Resources\UcmaRuntimeSetup.exe"
                Write-Host "Done!" -ForegroundColor "Yellow"
                Write-Host "Scanning for .NET Framework 4.5.2..."
            }
            ElseIf($UCMARuntimeSetup -eq $true) {
                $UCMARuntimeSetup = "C:\Exchange_Resources\UcmaRuntimeSetup.exe"
                Remove-Item $UCMARuntimeSetup -Force
                Clear-Host
                Write-Host "Downloading UCMA Runtime Setup..."
                Package-Download "http://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe" "C:\Exchange_Resources\UcmaRuntimeSetup.exe"
                Write-Host "Done!" -ForegroundColor "Yellow"
                Write-Host "Scanning for .NET Framework 4.5.2..."
            }
        }
    }

    Function Install-DOTNET {
        $DOTNETInstallCheck = '*Microsoft .NET Framework 4.5.2*'
        $DOTNETInstalled = Get-InstalledApps | where {$_.DisplayName -like $DOTNETInstallCheck}
        Clear-Host
        If ($DOTNETInstalled -eq $null) {
            Write-Host "Downloading .NET Framework 4.5.2..." `n
            $NDP452 = Test-Path "C:\Exchange_Resources\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
            If($NDP452 -eq $false) {
                Package-Download "http://download.microsoft.com/download/E/2/1/E21644B5-2DF2-47C2-91BD-63C560427900/NDP452-KB2901907-x86-x64-AllOS-ENU.exe" "C:\Exchange_Resources\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
                Write-Host "Done!" -ForegroundColor "Yellow" `n
                Write-Host "Installing .NET Framework 4.5.2..." `n
                Start-Process -FilePath "C:\Exchange_Resources\NDP452-KB2901907-x86-x64-AllOS-ENU.exe" -ArgumentList "/passive /norestart" -wait
                Write-Host "Done!" -ForegroundColor "Yellow" `n
            }
            ElseIf($NDP452 -eq $true) {
                $NDP452 = "C:\Exchange_Resources\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
                Remove-Item $NDP452 -Force
                Clear-Host
                Write-Host "Downloading .NET Framework 4.5.2..." `n
                Package-Download "http://download.microsoft.com/download/E/2/1/E21644B5-2DF2-47C2-91BD-63C560427900/NDP452-KB2901907-x86-x64-AllOS-ENU.exe" "C:\Exchange_Resources\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
                Write-Host "Done!" -ForegroundColor "Yellow" `n
                Write-Host "Installing .NET Framework 4.5.2..." `n
                Start-Process -FilePath "C:\Exchange_Resources\NDP452-KB2901907-x86-x64-AllOS-ENU.exe" -ArgumentList "/passive /norestart" -wait
                Write-Host "Done!" -ForegroundColor "Yellow" `n
            }
        }
        Else {
            Write-Host ".Net Framework 4.5.2 is installed!"
        }
    }
    Function Create-ExtractFolder {
        $TargetFolder = Test-Path "C:\Exchange_Resources\Exchange_2016_RTM_Extracted"
        If($TargetFolder) {
            $TargetFolder = "C:\Exchange_Resources\Exchange_2016_RTM_Extracted"
            Remove-Item $TargetFolder | Out-Null
            Write-Host `n "Extracting Exchange..." `n
            Start-Process -Filepath $ExchangeSetup -ArgumentList "/extract:$TargetFolder /u" -Wait
            Write-Host "Done!" -ForegroundColor "Yellow" `n
        }
        Else {
            $TargetFolder = "C:\Exchange_Resources\Exchange_2016_RTM_Extracted"
            Write-Host `n "Extracting Exchange..." `n
            Start-Process -Filepath $ExchangeSetup -ArgumentList "/extract:$TargetFolder /u" -Wait
            Write-Host "Done!" -ForegroundColor "Yellow" `n
        }
    }

    Function Reboots {
        # Specifies Startup Location for Install Script
        $Path = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
        Set-Content $Path {
                If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
                $Arguments = "& '" + $MyInvocation.mycommand.definition + "'"
                Start-Process powershell -Verb runAs -ArgumentList $arguments
                Break
            } 
            Remove-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
                        
            Start-Process -Filepath "C:\Exchange_Resources\UcmaRuntimeSetup.exe" -ArgumentList "/passive /norestart" -wait
            Start-Process -Filepath "C:\Exchange_Resources\filterpack2010sp2-x64-fullfile-en-us.exe" -ArgumentList "/passive /norestart" -wait
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
                $Organization = Get-Content -Path "C:\Exchange_Resources\OrganizationName.txt"
                $ExchangeInstallLocation = Get-Content -Path "C:\Exchange_Resources\InstallPath.txt" 
                Write-Host "Server is now ready to install Exchange Server 2016 RTM..."
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


   #####################################################################################
   #                                                                                   #
   #                               Exchange Options                                    #
   #                                                                                   #
   #                     Runs CLI parameters to install Exchange                       #              
   #                 Alter these parameters to your desired configuration              #
   #                                                                                   #
   #####################################################################################


                C:\Exchange_Resources\Exchange_2016_RTM_Extracted\setup.exe /PrepareSchema /IAcceptExchangeServerLicenseTerms
                C:\Exchange_Resources\Exchange_2016_RTM_Extracted\setup.exe /PrepareAD /OrganizationName:"$Organization" /IAcceptExchangeServerLicenseTerms                     
                C:\Exchange_Resources\Exchange_2016_RTM_Extracted\setup.exe /mode:Install /role:Mailbox /TargetDir:"$ExchangeInstallLocation" /OrganizationName:"$Organization" /IAcceptExchangeServerLicenseTerms       
                

                If (Get-Content C:\ExchangeSetupLogs\ExchangeSetup.log | Select-String -Pattern "The Exchange Server setup operation completed successfully.") {
                    Clear-Host
                    Write-Host `n `t "Exchange Install Complete!" -ForegroundColor Green `n
                }
                Else {
                    Clear-Host
                    Write-Host `n `t "Exchange Install did not complete. Please refer to the Exchange Setup Logs for troubleshooting." -ForegroundColor Red `n
                    Pause
                }
     # Changes file association for .ps1 to default (notepad.exe)
                Set-ItemProperty -Path HKLM:\SOFTWARE\Classes\Microsoft.PowerShellScript.1\Shell\Open\Command -Name "(Default)" -Value '"C:\Windows\System32\notepad.exe" "%1"' -Type String -Force
     # Changes QuickEdit to default registry setting
                Set-ItemProperty -Path HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe -Name "QuickEdit" -Value '1' -Type DWord -Force
     # Changes DoNotOpenServerManagerAtLogon to default registry setting
                New-ItemProperty -Path HKCU:\Software\Microsoft\ServerManager -Name "DoNotOpenServerManagerAtLogon" -Value "0" -Type DWord -Force
     # Removes InstallExchange.ps1 script from startup folder
                Remove-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp\InstallExchange.ps1"
     # Removes leftover text file for Organization Name
                Remove-Item "C:\Exchange_Resources" -Recurse -Force
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
    }

   #####################################################################################
   #                                                                                   #
   #                                 Main Script                                       #
   #                                                                                   #
   #####################################################################################


    # Mimics the console color scheme to match that of the Exchange Management Shell. Changes size of console window to support Welcome screen. 
    Set-Console
    Disable-FireWall
    # Disables User Account Control
    Disable-UAC
    # Disables Open File Security Warnings
    Disable-OpenFileSecurityWarning
    # Disables IE Security Warnings
    Disable-IEESC
    # Welcome Screen
    Display-WelcomeScreen
    Check-DomainMembership
    # Checks if C:\Exchange Resources already exists. If it does not, it creates the Exchange Resources Directory
    Create-ResourcesDirectory | Out-Null
    Set-AutomaticLogon
    $Downloaded_Title = "" 
    $Downloaded_Message = "Have you already downloaded Exchange 2016 RTM?"
    $Downloaded_Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
    $Downloaded_No = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
    $Downloaded_Options = [System.Management.Automation.Host.ChoiceDescription[]]($Downloaded_Yes, $Downloaded_No)
    $Downloaded = $host.ui.PromptForChoice($Downloaded_Title, $Downloaded_Message, $Downloaded_Options, 0) 
        switch ($Downloaded) {
            0 {
                Write-Host `n
                $Extracted_Title = "" 
                $Extracted_Message = "Has the Exchange 2016 RTM executable already been extracted?"
                $Extracted_Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
                $Extracted_No = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
                $Extracted_Options = [System.Management.Automation.Host.ChoiceDescription[]]($Extracted_Yes, $Extracted_No)
                $Extracted = $host.ui.PromptForChoice($Extracted_Title, $Extracted_Message, $Extracted_Options, 0)
                switch ($Extracted) {
                    0 {
                        :OuterLoop Do {
                            Write-Host `n "Please select the location where Exchange 2016 RTM is extracted"
                            $ExchangeExtracted = Read-FolderBrowserDialog -Message "Exchange 2016 RTM Extracted Directory"                                                      
                            If (![string]::IsNullOrEmpty($ExchangeExtracted)) { 
                                Write-Host "You selected the directory: $ExchangeExtracted" -Foreground "Yellow" `n
                                $ExchangeSetup = $ExchangeExtracted + "\Exchange2016-x64.exe"
                            }
                            Else { 
                                Clear-Host
                                Write-Host `n "You did not select a location. Please select the Exchange Executable file" -foreground "Red" `n
                                $ExchangeExtracted = $null    
                            }
                        }
                        While ($ExchangeExtracted = $null)
                        :OuterLoop Do {
                            Write-Host `n "Please select the location where you would like to have Exchange 2016 RTM Installed" `n
                            $ExchangeInstallLocation = Read-FolderBrowserDialog -Message "Exchange 2016 RTM Install Directory"
                            If (![string]::IsNullOrEmpty($ExchangeInstallLocation)) { 
                                Write-Host "You selected the directory: $ExchangeInstallLocation" -Foreground "Yellow" 
                            }
                            Else { 
                            Clear-Host
                            Write-Host `n "You did not select a location. Please select the Exchange Executable file" -foreground "Red" `n
                            $ExchangeInstallLocation = $null    
                            }
                        }
                        While ($ExchangeInstallLocation -eq $null)
                      }
                    1 {
                        :OuterLoop Do {
                            $ExchangeExtracted = Test-Path "C:\Exchange_Resources\Exchange_2016_RTM_Extracted"
                            If($ExchangeExtracted -eq $false) {
                                New-Item "C:\Exchange_Resources\Exchange_2016_RTM_Extracted" -type directory
                                $ExchangeExtracted = "C:\Exchange_Resources\Exchange_2016_RTM_Extracted"                                              
                            }
                            ElseIf($ExchangeExtracted -eq $true) { 
                                $ExchangeExtracted = "C:\Exchange_Resources\Exchange_2016_RTM_Extracted"
                            }
                        }
                        While ($ExchangeExtracted = $null) 
                        :OuterLoop Do {
                            Clear-Host
                            Write-Host `n "Please choose the directory where the download for Exchange 2016 RTM is located" `n
                            $ExchangeDownload = Read-FolderBrowserDialog -Message "Please select a directory"
                            If (![string]::IsNullOrEmpty($ExchangeDownload)) { 
                                Write-Host "You selected the directory: '$ExchangeDownload'" -ForegroundColor "Yellow" `n
                                $ExchangeSetup = $ExchangeDownload + "\Exchange2016-x64.exe"
                            }
                            Else {
                            Clear-Host
                            Write-Host `n "You did not select a location to download Exchange 2016 RTM. Please select a location." -foreground "Red" `n
                            $ExchangeDownload = $null                                     
                            }
                        }
                        While ($ExchangeDownload = $null)            
                        :OuterLoop Do {
                            Write-Host `n "Please select the location where you would like to have Exchange 2016 RTM Installed" `n
                            $ExchangeInstallLocation = Read-FolderBrowserDialog -Message "Exchange 2016 RTM Install Directory"
                            If (![string]::IsNullOrEmpty($ExchangeInstallLocation)) { 
                                Write-Host "You selected the directory: $ExchangeInstallLocation" -Foreground "Yellow"
                                Save-OrganizationName
                                $InstallPath = Test-Path "C:\Exchange_Resources\InstallPath.txt"
                                If($InstallPath) {
                                }
                                Else {
                                    New-Item "C:\Exchange_Resources\InstallPath.txt" -ItemType File | Out-Null
                                    Set-Content -Path "C:\Exchange_Resources\InstallPath.txt" -Value "$ExchangeInstallLocation"
                                }
                                Display-ReadyScreen 
                            }
                            Else { 
                                Clear-Host
                                Write-Host `n "You did not select a location. Please select the Exchange Executable file" -foreground "Red" `n
                                $ExchangeInstallLocation = $null    
                            }
                        }
                        While ($ExchangeInstallLocation -eq $null)
                                }            
                            }
                        }
            1 {
                :OuterLoop Do {
                    Write-Host `n "Please select the location where you would like to have Exchange 2016 RTM Installed" `n
                    $ExchangeInstallLocation = Read-FolderBrowserDialog -Message "Exchange 2016 RTM Install Directory"
                    If (![string]::IsNullOrEmpty($ExchangeInstallLocation)) { 
                        Write-Host "You selected the directory: $ExchangeInstallLocation" -Foreground "Yellow" 
                    }
                    Else { 
                        Clear-Host
                        Write-Host `n "You did not select a location. Please select the Exchange Executable file" -foreground "Red" `n
                        $ExchangeInstallLocation = $null    
                    }
                }
                While ($ExchangeInstallLocation -eq $null)
                :OuterLoop Do { 
                    Write-Host `n "Please choose a directory to store the download for Exchange 2016 RTM" `n
                    $ExchangeDownload = Read-FolderBrowserDialog -Message "Please select a directory"
                    If (![string]::IsNullOrEmpty($ExchangeDownload)) { 
                        Write-Host "You selected the directory: '$ExchangeDownload'" -ForegroundColor "Yellow" 
                        $ExchangeSetup = $ExchangeDownload + "\Exchange2016-x64.exe"
                        Save-OrganizationName
                        $InstallPath = Test-Path "C:\Exchange_Resources\InstallPath.txt"
                        If($InstallPath) {
                        }
                        Else {
                            New-Item "C:\Exchange_Resources\InstallPath.txt" -ItemType File | Out-Null
                            Set-Content -Path "C:\Exchange_Resources\InstallPath.txt" -Value "$ExchangeInstallLocation"
                        }
                        Display-ReadyScreen
                        Write-Host "Downloading Installation Packages..." `n 
                        Package-Download "https://download.microsoft.com/download/3/9/B/39B8DDA8-509C-4B9E-BCE9-4CD8CDC9A7DA/Exchange2016-x64.exe" "$ExchangeSetup" 
                        Write-Host "Finished downloading Exchange2016-x64.exe"
                    }
                    Else { 
                        Clear-Host
                        Write-Host `n "You did not select a directory." -foreground "Red" `n
                        $ExchangeDownload = $null
                    }
                }
                While ($ExchangeDownload = $null)
              }
         }
# Download request for 2010 FilterPack SP2
    Install-FilterPack2  
# Download request for UcmaRuntimeSetup
    Install-UCMA
# Download request for .NET Framework 4.5.2
    Install-DOTNET
# Installation of Windows Feature Prerequisites
    Clear-Host
    Write-Host "Installing Windows Features..."
    Install-WindowsFeature RSAT-ADDS | Out-Null
    Install-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation | Out-Null
    Write-Host "Done!" -ForegroundColor "Yellow"
    Clear-Host
    Create-ExtractFolder
    Reboots
    Restart-Computer
