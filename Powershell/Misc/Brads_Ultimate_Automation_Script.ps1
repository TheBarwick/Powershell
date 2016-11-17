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
Write-Host “`t`t`t1. User Tasks” -Fore Cyan
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
Write-Host “`t`tPlease select the User administration task you require`n” -Fore Green
Write-Host “`t`t`t1. Create New User” -Fore Green
Write-Host “`t`t`t2. Delete Existing User” -Fore Green
Write-Host “`t`t`t3. Rename Existing User” -Fore Green
Write-Host “`t`t`t4. Add a User as Local Admin” -Fore Green
Write-Host “`t`t`t5. Enroll VBT-User WiFi Certificate” -Fore Green
Write-Host “`t`t`t6. Go to Main Menu`n” -Fore Green
[int]$xMenu2 = Read-Host “`tEnter Menu Option Number”
if( $xMenu1 -lt 1 -or $xMenu1 -gt 6 ){
Write-Host “`tPlease select one of the options available.`n” -Fore Red;start-Sleep -Seconds 1
}
}
Switch ($xMenu2){
1{

Write-Host “`n`t##############################################################" -Fore Green
Write-Host "`t#                INCLUDE DOMAIN IN USERNAME                  #" -Fore Green
Write-Host "`t##############################################################" -Fore Green

Import-Module ActiveDirectory

# Prompts user to answer Yes or No to question specified in $caption
Function checkContinue($caption, $message, $default) {
	$yes = new-Object System.Management.Automation.Host.ChoiceDescription "&yes","Yes"
	$no = new-Object System.Management.Automation.Host.ChoiceDescription "&no","No"
	$choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)
	$answer = $host.ui.PromptForChoice($caption,$message,$choices,$default)

	switch ($answer){
		0 {return $True}
		1 {return $False}
	}
}

# Adds a share[d folder] at the specified $path and assigns proper permissions to $username
Function addFolder ($path, $username) {
    New-Item $path -type directory
	$acl = Get-Acl $path
	$perm = "voicebox\$username","Modify","ContainerInherit,ObjectInherit","None","Allow"
	$ar = New-Object system.security.accesscontrol.filesystemaccessrule $perm
	$acl.SetAccessRule($ar)
	$acl | Set-Acl $path
}

# Asks for domain admin credentials (include domain in username)
$cred = get-credential

Clear-Host

Write-Host "`nPlease enter the information for the new user below. First and Last name are required.`n"

# Prompts for name
$firstName = (Read-Host "First Name").trim()
$lastName = (Read-Host "Last Name").trim()

# Checks if $firstName or $lastName are blank
If (($firstName -eq "") -Or ($lastName -eq "")) {
	Do {
		Write-Host "`nFirst or Last name fields cannot be blank. Press any key to exit..."
		$firstName = (Read-Host "First Name").trim()
		$lastName = (Read-Host "Last Name").trim()
	} While (($firstName -eq "") -Or ($lastName -eq ""))
}

# Prompts for title, manager and department
$title = (Read-Host "Job Title").trim()
$manager = (Read-Host "Manager's Alias").trim()
$department = (Read-Host "Department").trim()

# Checks if manager is a valid user
If ((Get-ADUser -LDAPFilter "(SAMAccountName=$manager)") -eq $Null) {
	Do {
		Write-Host "`nManager $manager not found. Please re-enter the manager's alias."
		$manager = (Read-Host "Manager").trim()
	} While ((Get-ADUser -LDAPFilter "(SAMAccountName=$manager)") -eq $Null)
}

# Asks for seat location if known
#$office = (Read-Host "Office Location (ex. PP100_60 or MPW270_420. Leave blank if unknown.)").trim()

#$location = ""
#do {
#    $location = (Read-Host "Office Location (Available options: MPW, PP, Other, Unknown)").trim()
#    If ($location.ToLower() -eq "other") {
#        Write-Host "`n(!) You will need to manually fill in the 'Office' field in AD for this user" -foreground "red"
#        break
#    } ElseIf ($location.ToLower() -eq "Unknown") {
#        Write-Host "`nSkipping this question..."
#        break
#    } ElseIf ($location.ToLower() -ne "mpw" -or $location.ToLower() -eq "pp") {
#        Write-Host "Must be one of the options above. Try again."
#    }
#}
#while (($location.ToLower() -ne "mpw") -or ($location.ToLower() -ne "pp"))

#$seat = 0
#If (($office.ToLower() -ne "other" -or $office.ToLower() -ne "unknown")) {
#    Do {
#        $seat = (Read-Host "Seat Number").trim()
#        If ($seat -ne [int]) {
#        Write-Host "Not a valid number. Try again."
#        }
#    }
#    while ($seat -ne [int])
#}

#$office = $location.ToUpper() + "_" + $seat
#Write-Host "Using " + $office + "as the office location"

# Asks if user is Bellevue FTE (0), Bellevue contractor (1) or branch office employee (2)
$fte = 0
If (checkContinue "Is this user a *Full-Time* employee in Bellevue?" "Select 'N' for part-time/contractor or non-Bellevue office employees" 0) {
	Write-Host "`nUser is a full-time employee at the Bellevue office`n"
} Else {
	If (checkContinue "Is this user a *Part-Time or Contract* employee in Bellevue?" "Select 'N' for non-Bellevue office employees" 0) {
		Write-Host "`nUser is a contractor/part-time employee at the Bellevue office"
		$fte = 1
	} Else {
		Write-Host "`n(!) Please manually add the user into the appropriate branch office's group.`n" -foreground "red"
		$fte = 2
	}
}

# Checks if name exists in AD
Write-Host "`nChecking if name $firstName $lastName exists.......... " -NoNewLine
$checkNameExists = Get-ADUser -LDAPFilter "(Name=$firstName $lastName)"
If ($checkNameExists -eq $Null) {
	"Name does not exist in AD"
} Else {
	Write-Host "Name exists in AD"
	If (checkContinue "It is HIGHLY recommended to use an alternative name." "Do you want to continue?" 1) {Write-Host "Continuing..."}
	Else {exit}
}

# Checks if username exists in AD and generates alternatives if it can.
$username = ""
$position = 1
Write-Host "`nGenerating a new unique username.......... "

Do {
	If ($position -le $lastName.length) {
		$username = $firstName.toLower() + $lastName.toLower().Substring(0,$position)
		$position++
	} Else {
		Write-Host "`(!)Error: Cannot autogenerate username. Enter a custom username below."
		$username = (Read-Host 'Custom Username').trim()
	}
} While ((Get-ADUser -LDAPFilter "(SAMAccountName=$username)") -ne $Null)

# Asks if auto-generated username is OK. If not, ask for a custom user name
If (checkContinue "The generated username is '$username'." "Do you want to use this?" 0) {
	Write-Host "`nUsing username '$username'"
} Else {
	Write-Host "`nEnter a custom username below."
	$username = (Read-Host 'Custom Username').trim()
}

# Generates a password based on name
$passPlain = "%Temp" + $firstName.toUpper().Substring(0,1) + $lastName.toUpper().Substring(0,1) + "%"
Write-Host "`nPassword will be set to $passPlain"
$passSecure = ConvertTo-SecureString $passPlain -AsPlainText -Force

# Confirm account creation
If (checkContinue "This will create the new user's account." "Do you want to continue?" 1) {
	Write-Host "`nAdding new user $firstName $lastName`n"
} Else {
	exit
}

# Build the New-ADUser command
$titleCmd = "-Title '$title'"
$managerCmd = "-Manager '$manager'"
$departmentCmd = "-Department '$department'"
$newADUserCmd = "New-ADUser -Name '$firstName $lastName' -SAMAccountName $username -GivenName $firstName -Surname $lastName -DisplayName '$firstName $lastName' -ChangePasswordAtLogon `$True -AccountPassword `$passSecure -Company 'VoiceBox Technologies Inc.' $titleCmd $managerCmd $departmentCmd -Credential `$cred"

# Runs the built New-ADUser command
invoke-expression $newADUserCmd

# Adds to ChatUsers and BellevueEmployees/BellevueContractors groups
Write-Host "`nAdding to default groups`n"
Add-ADGroupMember -identity "ChatUsers" -member $username -Credential $cred
If ($fte -eq 0) {Add-ADGroupMember -identity "Bellevue Employees" -member $username -Credential $cred}
ElseIf ($fte -eq 1) {Add-ADGroupMember -identity "Bellevue Contractors" -member $username -Credential $cred}
Else {Write-Host "(!) Remember to add user to the appropriate branch office group" -foreground "red"}

Write-Host "`nCreating mailbox for $username`n"

# Opens a remote PS session on Mail02
$mailSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mail02/PowerShell/ -Authentication Kerberos -Credential $cred
Import-PSSession $mailSession

# Creates and links a mailbox on the newly created AD user
Enable-Mailbox -Identity "$username" -Alias "$username" -DisplayName "$firstName $lastName"

Remove-PSSession $mailSession

# Open a remote PS session on Filer
$filerSession = New-PSSession filer -Authentication Kerberos -Credential $cred
Enter-PSSession $filerSession

# Creates User share folder on Filer. These three lines can be copied to auto-create additional shares.
Write-Host "`nCreating User share for $username`n"
$userPath = "D:\Users\$username"
Invoke-Command -Session $filerSession -ScriptBlock ${function:addFolder} -ArgumentList $userPath,$username		# uses addFolder function to do the actual work

Exit-PSSession
Remove-PSSession $filerSession

Write-Host "`nUser created successfully. Press any key to exit..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

}
2{

Write-Host “`n`t##############################################################" -Fore Green
Write-Host "`t#                INCLUDE DOMAIN IN USERNAME                  #" -Fore Green
Write-Host "`t##############################################################" -Fore Green

$a = (Get-Host).UI.RawUI
$a.BackgroundColor = "black"
$a.ForegroundColor = "magenta"

# Asks for domain admin credentials (include domain in username)

Import-Module ActiveDirectory

$cred = get-credential

Clear-Host

$username = (Read-Host Username).trim()

$groupname = (Read-Host Group Name).trim()

Remove-ADUser -Identity $username -Credential $cred

}
3{ 

Write-Host “`n`t##############################################################" -Fore Green
Write-Host "`t#                INCLUDE DOMAIN IN USERNAME                  #" -Fore Green
Write-Host "`t##############################################################" -Fore Green

$a = (Get-Host).UI.RawUI
$a.BackgroundColor = "black"
$a.ForegroundColor = "magenta"

# Asks for domain admin credentials (include domain in username)

Import-Module ActiveDirectory

$cred = get-credential

Clear-Host

$username = (Read-Host Username).trim()

$newname = (Read-Host New Name).trim()

Rename-ADObject -Identity $username -NewName $newname -Credential $cred

}
4{ 
$a = (Get-Host).UI.RawUI
$a.BackgroundColor = "black"
$a.ForegroundColor = "green"

$user = (Read-host "`n`n`n`nEnter your Admin Username")
runas /user:voicebox\$user "powershell \\filer\documents\IT\_Scripts\Add_a_User_as_Local_Admin.ps1" 

}
5{ certreq -enroll -q user_cert_autoenroll}
6{ Write-Host “`n`tYou Selected Option 5 – Quit the Administration Tasks`n” -Fore Yellow; break}
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


$cred = get-credential

Clear-Host
$groupname = (Read-Host Distribution Group Name)
$mailSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mail02/PowerShell/ -Authentication Kerberos -Credential $cred
Import-PSSession $mailSession

New-DistributionGroup -Name "$groupname" -OrganizationalUnit "voiceboxtechnologies.com/Engineering" 
Remove-PSSession $mailSession
Clear-Host
Import-Module ActiveDirectory

Write-Host “`t`t Hit Enter Twice Once All Members Have Been Specified`n” -Fore Green
Write-Host “`t`t Members = Domain User Name`n” -Fore Green

Add-ADGroupMember -Identity $groupname -Credential $cred



}
2{ 

Write-Host “`n`t##############################################################" -Fore Green
Write-Host "`t#                INCLUDE DOMAIN IN USERNAME                  #" -Fore Green
Write-Host "`t##############################################################" -Fore Green

$a = (Get-Host).UI.RawUI
$a.BackgroundColor = "black"
$a.ForegroundColor = "magenta"


Import-Module ActiveDirectory

$cred = get-credential

Clear-Host


Write-Host "Members = User Name`n"


$groupname = (Read-Host Group Name)

Add-ADGroupMember -Identity $groupname -credential $cred

}

3{ 

Write-Host “`n`t##############################################################" -Fore Green
Write-Host "`t#                INCLUDE DOMAIN IN USERNAME                  #" -Fore Green
Write-Host "`t##############################################################" -Fore Green

$a = (Get-Host).UI.RawUI
$a.BackgroundColor = "black"
$a.ForegroundColor = "magenta"

$cred = get-credential

Clear-Host
$user = (Read-Host User Name to Remove)
$groupname = (Read-Host Distribution Group Name)
$mailSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mail02/PowerShell/ -Authentication Kerberos -Credential $cred
Import-PSSession $mailSession

Remove-DistributionGroupMember -Identity "$groupname" -Member "$user" 
Remove-PSSession $mailSession
Clear-Host

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

$cred = get-credential

Clear-Host

$username = (Read-Host Mailbox Name)

$mailSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mail02/PowerShell/ -Authentication Kerberos -Credential $cred

Import-PSSession $mailSession

Enable-Mailbox -Identity "$username" -Alias "$username" -DisplayName "$username"

Remove-PSSession $mailSession

}
2{ 

Write-Host “`n`t##############################################################" -Fore Green
Write-Host "`t#                INCLUDE DOMAIN IN USERNAME                  #" -Fore Green
Write-Host "`t##############################################################" -Fore Green

$a = (Get-Host).UI.RawUI
$a.BackgroundColor = "black"
$a.ForegroundColor = "magenta"

$cred = get-credential

Clear-Host

$mailbox = (Read-Host Mailbox Name)

$mailSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mail02/PowerShell/ -Authentication Kerberos -Credential $cred

Import-PSSession $mailSession

Remove-Mailbox -Identity voicebox\"$mailbox" 

Remove-PSSession $mailSession

 }
3{ 

Write-Host “`n`t##############################################################" -Fore Green
Write-Host "`t#                INCLUDE DOMAIN IN USERNAME                  #" -Fore Green
Write-Host "`t##############################################################" -Fore Green

$a = (Get-Host).UI.RawUI
$a.BackgroundColor = "black"
$a.ForegroundColor = "magenta"

$cred = get-credential

Clear-Host

$mailcontact = (Read-Host Mail Contact Name)
$alias = (Read-Host Contact Alias)
$externalemail = (Read-Host External E-mail Address)

$mailSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mail02/PowerShell/ -Authentication Kerberos -Credential $cred

Import-PSSession $mailSession

New-MailContact -Name "$mailcontact" -ExternalEmailAddress "$external" -Alias "$alias"

Remove-PSSession $mailSession

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
\\filer\documents\IT\_Scripts\Brads_Ultimate_Automation_Script    #… Loop the function
}