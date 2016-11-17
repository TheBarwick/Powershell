Unattended Installation of Exchange 2013
=====
Preface:
-----------
There are many other iterations of automated unattended installations of Exchange available out on the internet. The biggest problem I ran into was that they were either too complex, or don’t work without a series of pre-requisites etc. I wanted to build out a script that would ask me the important stuff, specifically what I would need to input as the administrator that would effectively install Exchange into an environment. At the same time, I wanted to see the progress of my downloads, and installations. Therefore, I have taken bits and pieces of code from these other versions as well as created my own functions, processes, and overall experience to accomplish a smooth, automated installation of Exchange 2013. This script is specifically tailored for Exchange 2013 CU11, and is easily tailored to meet the needs of previous and upcoming CU’s as well as Exchange 2016. I plan to continue to enhance and improve the scripts performance and ability in the near future.

Title:
-----------
Unattended_Exchange2013_CU11_Installation.ps1

Requirements:
-----------
* [PowerShell 3](http://www.microsoft.com/en-us/download/details.aspx?id=34595)
* Windows Server 2008 R2 SP1, 2012, 2012 R2;
* Domain-joined system;
* Domain Account with Domain Admins, Enterprise Admins, and Schema Admins membership;
* System is in same site as Schema Master;
* PowerShell v3
* 15GB of storage space on C: drive available

Description:
-----------
* This script will install Exchange 2013 CU11 prerequisites, create the Exchange Organization (prepares Active Directory) and installs Exchange Server.
* Currently set to install a multi-role server of Exchange.

Script Assumptions:
-----------
* None of the pre-requisites have been downloaded or installed. 
* The user running the script has proper permissions such as domain group membership required by Exchange.
* This is the first Exchange installation in a new environment. 
  * The script can be ran to implement more Exchange servers if desired, but does no checks of previous installations of Exchange such as Exchange 2007, 2010, etc. 
* The server the script is running on is in the same Active Directory site as the Schema Master.
* 3rd Party Firewalls have been disabled

High Level List of Features
-----------
* Automatically runs PowerShell console with Administrator privileges if applicable
  * Checks for Domain Membership. 
* If the script is ran on a server that is not domain joined, script will prompt user and then exit.
* Disables/Re-Enables UAC
* Disables/Re-Enables IE Enhanced Security Configuration
* Disables/Re-Enables File Security Warning Dialogs
* Provides real-time progress of installations and configurations
* Upon reboot, script resumes where it left off after an automatic logon using credentials provided at script pre-requisite prompt.
  * Script will store credentials provided in the registry of the server that is running the script. Once the script has finished it will delete the registry key attributes holding these credentials. 
* Downloads all packages required as part of the Exchange Installation process such as: 
  * Remote Tools Administration Pack
  * Required Windows Roles and Features
  * .NET Framework 4.5.2
  * Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit
* Prompts for download and installation directory specifications outside default install location.

High Level Order of Operations:
-----------
1.	Disables Windows Firewall
2.	Disables UAC
3.	Disables IE Enhanced Security Configuration
4.	Disables File Security Warning Dialogs
5.	Disables PowerShell Console Quick Edit Mode
6.	Disables Automatic Server Manager Open at Logon
7.	Alters Default .File Association for .ps1 files from Notepad.exe to PowerShell.exe 
8.	Prompts Welcome Screen
9.	Check for Domain Membership
10.	Prompts for credentials
11.	Sets Automatic Logon
12.	Creates Temporary Exchange Resource Directory on C: Drive
13.	Prompts for download/installation/extraction locations
14.	Creates Extract Folder
15.	Creates Startup Script for resume on 1st reboot
16.	Installs .Net
17.	Installs Windows Features
18.	Reboots
19.	Resumes and Installs UCMA, FilterPack2
20.	Reboots
21.	Installs Exchange
22.	Removes Automatic Logon
23.	Re-Enables PowerShell Console Quick Edit Mode
24.	Re-Enables Automatic Server Manager Open at Logon
25.	Reverts default file association for .ps1 back to Notepad.exe 
26.	Deletes Temporary Exchange Resources Directory on C: Drive
27.	Re-Enables UAC
28.	Re-Enables IE Enhanced Security Configuration
29.	Re-Enables Windows Firewall
30.	Re-Enables Open File Security Warnings
31.	Restarts Server

Inspiration
-----------

* [Exchange v15 (2013/2016) Unattended Installation Script](https://gallery.technet.microsoft.com/office/Exchange-2013-Unattended-e97ccda4)
