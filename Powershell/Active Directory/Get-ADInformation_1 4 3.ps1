###########################################
# Get-ADInformation.ps1 v1.4.3
#
# Script by: Elan Shudnow and Jason Anderson
# Last Updated: 05/16/2014
########### Supported Platforms ############
# Active Directory: All Versions
#
# Note: This script requires all pre-Windows 2008 R2 DCs to have the "Active Directory Module for Windows PowerShell installed." If there are no 2008 R2 or above
# 		DCs, you will also need at least one DC with the "Active Directory Management Gateway Service" installed. Running this script without meeting these requirements
#		will still cause the script to run successfully.  However, not all information will be captured.
########### Permissions Required ###########
# Single Domain: Domain Admins
# Multi-Domain: Enterprise Admins
############### Misc Required ##############
# Run on a machine that has PowerShell and the ActiveDirectory Module available
################# Features ################# 
# To see a list of features, refer back to the CDW e-mail that contains Feature List and ChangeLog
############################################

#Set this to the directory you want to output script information to
$Directory = "C:\ADInfo\Get-ADInfo"

###########################################
###### Do Not Modify Below This Line ######
###########################################

$erroractionpreference = "SilentlyContinue"

function GenerateForm {

[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

$form1 = New-Object System.Windows.Forms.Form
$button1 = New-Object System.Windows.Forms.Button
$button2 = New-Object System.Windows.Forms.Button
$labelBox1 = New-Object System.Windows.Forms.Label
$labelBox2 = New-Object System.Windows.Forms.Label
$labelBox3 = New-Object System.Windows.Forms.Label
$checkBox3 = New-Object System.Windows.Forms.CheckBox
$checkBox2 = New-Object System.Windows.Forms.CheckBox
$checkBox1 = New-Object System.Windows.Forms.CheckBox
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
$labelBox1.Text = "User Info: Will increase script processing time by obtaining all user information across the forest"
$labelBox2.Text = "Computer Info: Will increase script processing time by obtaining all computer information across the forest"
$labelBox3.Text = "DC Info: Will increase script processing time by obtaining all DC event logs, DC Diag, etc. across the forest.  OS/Hardware/IP will still be gathered."
$labelBox3.AutoSize = $true

$b1= $false
$b2= $false
$b3= $false

#----------------------------------------------
#Generated Event Script Blocks
#----------------------------------------------

# Run Script Button Code When Button is Clicked
$handler_button1_Click= 
{
    #$labelBox1.Items.Clear();    

    if ($checkBox1.Checked)     {  
		#$labelBox1.Items.Add( "User Information will be gathered. This option adds a significant amount of time to script processing."  ) 
		$Global:GetUserInfo = $True
	}

    if ($checkBox2.Checked)    {  
		#$labelBox1.Items.Add( "Computer Information will be gathered. This option adds a significant amount of time to script processing."  )
		$Global:GetComputerInfo = $True
	}

    if ($checkBox3.Checked)    {  
		#$labelBox1.Items.Add( "Domain Controller Information will be gathered."  ) 
		$Global:GetDCInfo = $True
	}
	$form1.Close()
    #if ( !$checkBox1.Checked -and !$checkBox2.Checked -and !$checkBox3.Checked ) {   $labelBox1.Items.Add("No CheckBox selected....")}
	
}

# Cancel Button Code When Button is Clicked
$handler_button2_Click= 
{
	$form1.Close()
	$global:CancelScript = $true
}

#----------------------------------------------
# Building the main form
#region Generated Form Code
$form1.Text = "Get-ADInformation"
$form1.Name = "form1"
$form1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 1035
$System_Drawing_Size.Height = 167
$form1.ClientSize = $System_Drawing_Size

# Building the Run Script button
$button1.TabIndex = 6
$button1.Name = "button1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 23
$button1.Size = $System_Drawing_Size
$button1.UseVisualStyleBackColor = $True

$button1.Text = "Run Script"

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 27
$System_Drawing_Point.Y = 128
$button1.Location = $System_Drawing_Point
$button1.DataBindings.DefaultDataSourceUpdateMode = 0
$button1.add_Click($handler_button1_Click)

$form1.Controls.Add($button1)

# Building the Cancel Script button
$button2.TabIndex = 7
$button2.Name = "button2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 85
$System_Drawing_Size.Height = 23
$button2.Size = $System_Drawing_Size
$button2.UseVisualStyleBackColor = $True

$button2.Text = "Cancel Script"

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 110
$System_Drawing_Point.Y = 128
$button2.Location = $System_Drawing_Point
$button2.DataBindings.DefaultDataSourceUpdateMode = 0
$button2.add_Click($handler_button2_Click)

$form1.Controls.Add($button2)

#Building the Label for GetUserInfo
#$labelBox1.FormattingEnabled = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 715
$System_Drawing_Size.Height = 15
$labelBox1.Size = $System_Drawing_Size
$labelBox1.DataBindings.DefaultDataSourceUpdateMode = 0
$labelBox1.Name = "labelBox1"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 167
$System_Drawing_Point.Y = 25
$labelBox1.Location = $System_Drawing_Point
$labelBox1.TabIndex = 5

$form1.Controls.Add($labelBox1)

# Building the Label for GetComputerInfo
#$labelBox2.FormattingEnabled = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 715
$System_Drawing_Size.Height = 15
$labelBox2.Size = $System_Drawing_Size
$labelBox2.DataBindings.DefaultDataSourceUpdateMode = 0
$labelBox2.Name = "labelBox2"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 167
$System_Drawing_Point.Y = 58
$labelBox2.Location = $System_Drawing_Point
$labelBox2.TabIndex = 4

$form1.Controls.Add($labelBox2)

# Building the Label for GetDCInfo
#$labelBox3.FormattingEnabled = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 715
$System_Drawing_Size.Height = 15
$labelBox3.Size = $System_Drawing_Size
$labelBox3.DataBindings.DefaultDataSourceUpdateMode = 0
$labelBox3.Name = "labelBox3"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 167
$System_Drawing_Point.Y = 88
$labelBox3.Location = $System_Drawing_Point
$labelBox3.TabIndex = 3

$form1.Controls.Add($labelBox3)

# Building the CheckBox for GetDCInfo
$checkBox3.UseVisualStyleBackColor = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 104
$System_Drawing_Size.Height = 24
$checkBox3.Size = $System_Drawing_Size
$checkBox3.TabIndex = 2
$checkBox3.Text = "DCInfo"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 27
$System_Drawing_Point.Y = 85
$checkBox3.Location = $System_Drawing_Point
$checkBox3.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox3.Name = "checkBox3"

$form1.Controls.Add($checkBox3)

# Building the CheckBox for GetComputerInfo
$checkBox2.UseVisualStyleBackColor = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 124
$System_Drawing_Size.Height = 24
$checkBox2.Size = $System_Drawing_Size
$checkBox2.TabIndex = 1
$checkBox2.Text = "ComputerInfo"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 27
$System_Drawing_Point.Y = 55
$checkBox2.Location = $System_Drawing_Point
$checkBox2.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox2.Name = "checkBox2"

$form1.Controls.Add($checkBox2)


# Building the CheckBox for GetUserInfo
$checkBox1.UseVisualStyleBackColor = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 104
$System_Drawing_Size.Height = 24
$checkBox1.Size = $System_Drawing_Size
$checkBox1.TabIndex = 0
$checkBox1.Text = "UserInfo"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 27
$System_Drawing_Point.Y = 23
$checkBox1.Location = $System_Drawing_Point
$checkBox1.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox1.Name = "checkBox1"

$form1.Controls.Add($checkBox1)

 # create Cancel button
  $CancelButton = New-Object System.Windows.Forms.Button
  $CancelButton.Location = New-Object System.Drawing.Size(160,320)
  $CancelButton.Size = New-Object System.Drawing.Size(75,23)
  $CancelButton.Text = "Cancel"
  $CancelButton.Add_Click(
  {$Looping=$False
   $RestoreFromFileForm.Close()
   exit
  })
  $RestoreFromFileForm.Controls.Add($CancelButton)


#Save the initial state of the form
$InitialFormWindowState = $form1.WindowState
#Init the OnLoad event to correct the initial state of the form
$form1.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$form1.ShowDialog()| Out-Null

} #End Function

#Call the Function
GenerateForm

if ($global:CancelScript) { 
	Write-Host "Terminating Script" -ForegroundColor Red
	$Global:CancelScript = $null
	exit
}

if (! (Get-Module -name "ActiveDirectory" -ea 0)){import-module activedirectory}

#Build up Unique Foldername and create required path
$filepath = $directory + "\{0:yyyyMMdd-HHmm}\" -f (Get-Date)
if(!(Test-Path -Path $filepath)) {New-Item -Path $filepath -Type directory | Out-Null}

Write-Host " "
Write-Host -ForegroundColor Yellow "INFO: Running Script and Outputting Data to" $Directory".  In a multi-domain environment, it is required run this script with Enterprise Admins access for proper WMI access."
Write-Host " "

# Directory where files will be saved
$useroutdir = $filepath + "\" + "UserInfo"
$computeroutdir = $filepath + "\" + "ComputerInfo"
$dcrootdir = $filepath + "\" + "Domain Controllers"
$adoutdir = $filepath + "\" + "ActiveDirectory"

# Function for when the script cycles through every DC to make sure it's pingable before gathering information.
function Ping-Host ($Server) {
       $result = Gwmi -Query "SELECT * FROM Win32_PingStatus WHERE Address='$Server'"
       if ($result.statuscode -eq 0) {$true} else {$false}
}

# create the new directory if it's not already there
if(!(test-path $useroutdir)){ mkdir $useroutdir | Out-Null }
if(!(test-path $computeroutdir)){ mkdir $computeroutdir | Out-Null }
if(!(test-path $dcrootdir)){ mkdir $dcrootdir | Out-Null }
if(!(test-path $adoutdir)){ mkdir $adoutdir | Out-Null }


# Return Active Directory Information (AD Schema and Exchange Schema)
Function Get-ADForestInformation {
	# Get some specific forest and domain info
	$objRoot=[adsi]"LDAP://rootDSE" 
	$objExchangePath="LDAP://CN=ms-Exch-Schema-Version-Pt,"+$objRoot.schemaNamingContext 
	$ExchangeSchema = [ADSI]"$objExchangePath"
	$objLyncPath="LDAP://CN=ms-RTC-SIP-SchemaVersion,"+$objRoot.schemaNamingContext 
	$LyncSchema = [ADSI]"$objLyncPath"
	$objADPath="LDAP://"+$objRoot.schemaNamingContext
	$ADSchema = [ADSI]"$objADPath"
	$Forest = [DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
	$Domains = $forest.Domains
	Write-Host ""
	Write-Host "- Starting to collect Forest-Wide Active Directory Information" -ForegroundColor Green
	Write-Host ""
	"Date Created: " + $(Get-Date -Format g)
	""
	"================= Forest Information ================="	
	# Obtain Exchange Schema Version
	if ($ExchangeSchema.rangeupper -eq 4397) {
		$ExSchemaVersion = "Version (4397): Exchange 2000 RTM"
	}
	elseif ($ExchangeSchema.rangeupper -eq 4406) {
		$ExSchemaVersion = "Version (4406): Exchange 2000 SP3"
	}
	elseif ($ExchangeSchema.rangeupper -eq 6870) {
		$ExSchemaVersion = "Version (6870): Exchange 2003 RTM"
	}
	elseif ($ExchangeSchema.rangeupper -eq 6936) {
		$ExSchemaVersion = "Version (6936): Exchange 2003 SP2"
	}
	elseif ($ExchangeSchema.rangeupper -eq 10628) {
		$ExSchemaVersion = "Version (10628): Exchange 2007 RTM"
	}
	elseif ($ExchangeSchema.rangeupper -eq 11116) {
		$ExSchemaVersion = "Version (11116): Exchange 2007 SP1"
	}
	elseif ($ExchangeSchema.rangeupper -eq 14622) {
		$ExSchemaVersion = "Version (14622): Exchange 2007 SP2 and Exchange 2010 RTM"
	}
	elseif ($ExchangeSchema.rangeupper -eq 14726) {
		$ExSchemaVersion = "Version (14711): Exchange 2010 SP1"
	}
	elseif ($ExchangeSchema.rangeupper -eq 14732) {
		$ExSchemaVersion = "Version (14732): Exchange 2010 SP2"
	}
	elseif ($ExchangeSchema.rangeupper -eq 14734) {
		$ExSchemaVersion = "Version (14734): Exchange 2010 SP3"
	}
	elseif ($ExchangeSchema.rangeupper -eq 15132) {
		$ExSchemaVersion = "Version (15132): Exchange 2013 Preview"
	}
	elseif ($ExchangeSchema.rangeupper -eq 15137) {
		$ExSchemaVersion = "Version (15137): Exchange 2013 RTM"
	}
	elseif ($ExchangeSchema.rangeupper -eq 15283) {
		$ExSchemaVersion = "Version (15283): Exchange 2013 SP1"
	}	
	 #Obtain LCS/OCS/Lync Schema Version
	if ($LyncSchema.rangeupper -eq 1006) {
		$LyncSchemaVersion = "Version (1006): LCS 2005"
	}
	elseif ($LyncSchema.rangeupper -eq 1007) {
		$LyncSchemaVersion = "Version (1007): OCS 2007 R1"
	}
	elseif ($LyncSchema.rangeupper -eq 1008) {
		$LyncSchemaVersion = "Version (1008): OCS 2007 R2"
	}
	elseif ($LyncSchema.rangeupper -eq 1100) {
		$LyncSchemaVersion = "Version (1100): Lync 2010"
	}
	elseif ($LyncSchema.rangeupper -eq 1150) {
		$LyncSchemaVersion = "Version (1150): Lync 2013"
	}

	# Obtain AD Schema Version
	if ($ADSchema.objectVersion -eq 13) {
		$ADSchemaVersion = "Version (13): Windows 2000"
	}
	elseif ($ADSchema.objectVersion -eq 30) {
		$ADSchemaVersion = "Version (30): Windows 2003 RTM/SP1/SP2"
	}
	elseif ($ADSchema.objectVersion -eq 31) {
		$ADSchemaVersion = "Version (31): Windows 2003 R2"
	}
	elseif ($ADSchema.objectVersion -eq 44) {
		$ADSchemaVersion = "Version (44): Windows 2008 RTM"
	}
	elseif ($ADSchema.objectVersion -eq 47) {
		$ADSchemaVersion = "Version (47): Windows 2008 R2 RTM"
	}
	elseif ($ADSchema.objectVersion -eq 52) {
		$ADSchemaVersion = "Version (52): Windows Server 8 Beta"
	}
	elseif ($ADSchema.objectVersion -eq 56) {
		$ADSchemaVersion = "Version (56): Windows Server 2012 RTM"
	}
	elseif ($ADSchema.objectVersion -eq 69) {
		$ADSchemaVersion = "Version (69): Windows Server 2012 R2"
	}

	# Gets Forest Information
	$Forest = [DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
	"Forest FQDN:" + $forest.Name
	"Forest Functional Level: " + $Forest.ForestMode
	"Domain Naming Master: " + $Forest.NamingRoleOwner
	"Schema Master: " + $Forest.SchemaRoleOwner
	"AD Schema: " + $ADSchemaVersion
	"Exchange Schema: " + $ExSchemaVersion
	"Lync Schema: " + $LyncSchemaVersion
}
Get-ADForestInformation | Out-File $adoutdir\ForestInformation.txt

Function Get-ADDomainInformation {
	# Get some specific forest and domain info
	$objRoot=[adsi]"LDAP://rootDSE" 
	$objExchangePath="LDAP://CN=ms-Exch-Schema-Version-Pt,"+$objRoot.schemaNamingContext 
	$ExchangeSchema = [ADSI]"$objExchangePath"
	$objLyncPath="LDAP://CN=ms-RTC-SIP-SchemaVersion,"+$objRoot.schemaNamingContext 
	$LyncSchema = [ADSI]"$objLyncPath"
	$objADPath="LDAP://"+$objRoot.schemaNamingContext
	$ADSchema = [ADSI]"$objADPath"
	$Forest = [DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
	$Domains = $forest.Domains
	"Date Created: " + $(Get-Date -Format g)
	""
	
	# Specifying Variables here to ensure that when the foreach loop for each domain runs, the existing variable does not get wiped
	# when it moves onto the next domain
	$ActiveComputersList = @()
	$AllComputersList = @()
	$AllComputersList = @()
	$UserSecurity = @()
	$DisabledComputerList = @()
	$InactiveComputerList = @()
	$ActiveServersList = @()
	$AdminList = @()
	$SAdminList = @()
	$EEAdminList = @()
	$DomainAdminList = @()
	
	# Gets each Domain in Forest and gets information about each Domain
	foreach ($Domain in $Domains) {
		$DomainName = $Domain.Name
		$DomainDN = (Get-ADDomain -Identity $DomainName).DistinguishedName
		"================= Domain Information ================="
		"Domain: " + $Domain.Name
		"Domain Functional Level: " + $Domain.DomainMode
		"PDC Emulator: " + $Domain.PdcRoleOwner
		"RID Master: " + $Domain.RidRoleOwner
		"Infrastructure Master: " + $Domain.InfrastructureRoleOwner
		$DomainControllers = $Domain.DomainControllers
		
		# get other relevant AD info		
		$ADPasswordPolicy = Get-ADDefaultDomainPasswordPolicy -server $DomainName 
		$UserobjPass = New-Object PSObject
		$UserobjPass | Add-Member NoteProperty "Domain" $DomainName
		$UserobjPass | Add-Member NoteProperty "ComplexityEnabled" $ADPasswordPolicy.ComplexityEnabled
		$UserobjPass | Add-Member NoteProperty "LockoutDuration" $ADPasswordPolicy.LockoutDuration
		$UserobjPass | Add-Member NoteProperty "LockoutObservationWindow" $ADPasswordPolicy.LockoutObservationWindow
		$UserobjPass | Add-Member NoteProperty "LockoutThreshold" $ADPasswordPolicy.LockoutThreshold
		$UserobjPass | Add-Member NoteProperty "MaxPasswordAge" $ADPasswordPolicy.MaxPasswordAge
		$UserobjPass | Add-Member NoteProperty "MinPasswordAge" $ADPasswordPolicy.MinPasswordAge
		$UserobjPass | Add-Member NoteProperty "MinPasswordLength" $ADPasswordPolicy.MinPasswordLength
		$UserobjPass | Add-Member NoteProperty "PasswordHistoryCount" $ADPasswordPolicy.PasswordHistoryCount
		$PasswordPolicies += $UserobjPass

		# Find all domain admins in current domain and append to CSV.  The final CSV will include Domain Admins from entire forest
		# and the first column will specify the Domain they are a Domain Admin for.
		$dadmins = (Get-ADGroup -Identity "Domain Admins" -Partition $DomainDN -properties member) 
		$dadminsmembers = $dadmins.member
		foreach ($DomainAdmin in $dadminsmembers) {
			$UserobjDadmins = New-Object PSObject
			$UserobjDadmins | Add-Member NoteProperty "Domain" $DomainName
			$UserobjDadmins | Add-Member NoteProperty "Member" $DomainAdmin
			$DomainAdminList += $UserobjDadmins
		}
		
		# Find all enterprise admins and append to CSV.. 
		$eadmins = (Get-ADGroup -Identity "Enterprise Admins" -Partition $DomainDN -properties member) 
		$eadminsmembers = $eadmins.member
		foreach ($EEAdmin in $eadminsmembers) {
			$UserobjEadmins = New-Object PSObject
			$UserobjEadmins | Add-Member NoteProperty "Domain" $DomainName
			$UserobjEadmins | Add-Member NoteProperty "Member" $EEAdmin
			$EEAdminList += $UserobjEadmins
		}
	
		# Find all schema admins and append to CSV.. 
		$sadmins = (Get-ADGroup -Identity "Schema Admins" -Partition $DomainDN -properties member) 
		$sadminsmembers = $sadmins.member
		foreach ($SAdmin in $sadminsmembers) {
			$UserobjSadmins = New-Object PSObject
			$UserobjSadmins | Add-Member NoteProperty "Domain" $DomainName
			$UserobjSadmins | Add-Member NoteProperty "Member" $SAdmin
			$SAdminList += $UserobjSadmins
		}

		# Find all administrators and append to CSV.. 
		$admins = (Get-ADGroup -Identity "Administrators" -Partition $DomainDN -properties member) 
		$adminsmembers = $admins.member

		foreach ($Admin in $adminsmembers) {
			$Userobjadmins = New-Object PSObject
			$Userobjadmins | Add-Member NoteProperty "Domain" $DomainName
			$Userobjadmins | Add-Member NoteProperty "Member" $Admin
			$AdminList += $Userobjadmins
		}
		
		""
		if ($GetDCInfo -eq $true) { 
			Write-Host " "
			Write-Host "- Starting to collect Domain Controller Information from $DomainName.  The script will collect OS, Hardware, IP information.  Because GetDCInfo was checked upon script launch, the following information will also be collected via Remote PowerShell: Event Log Information, systeminfo, dcdiag, netsh, dnscmd, w32tm, repadmin, netstatistics, AD Trusts, AD Site Links, Missing Subnets from AD Sites, AD Fine Grained Password Policy, and DNS info on DCs that have DNS installed." -ForegroundColor Green 
			Write-Host " "
		}
		else { 
			Write-Host " "
			Write-Host "- Starting to collect Domain Controller Information from $DomainName.  Because GetDCInfo was unchecked upon script launch, only OS, Hardware, and IP Information will be collected." -ForegroundColor Green 
			Write-Host " "
		}

		foreach ($DC in $DomainControllers) {
			$Ping = Ping-Host $DC.IPAddress
			$DCName = $DC.Name
			$DCShortName = $DCName.split(".")[0]
			$DCPS = Get-ADComputer $DCShortName -Properties operatingsystem
			Write-Host " "
			Write-Host "- Starting to collect Domain Controller Information from $DCName." -ForegroundColor Green
			Write-Host " "
			"Domain Controller: " + $DC.Name
			" > SiteName: " + $DC.Sitename
			" > IP: " + $DC.IPAddress
			if ($Ping -eq $True) { 
				" > Ping: Successful"
			
				# Adding -not $DCPS in case Active Directory Web Services are unavailable or not installed on all pre-Windows 2008 R2 DCs
				if (($DCPS.operatingsystem -notmatch '2012') -or (-not($DCPS))) {
				     $os = Get-WmiObject -class win32_operatingsystem -ComputerName $DCName 
				     $processor = Get-WmiObject -Class win32_Processor -ComputerName $DCName 
				     $memory = Get-WmiObject -Class win32_physicalmemory -ComputerName $DCName 
				     $volume = Get-WmiObject -Class win32_volume -ComputerName $DCName

			        " > OS Architecture: " + $os.OSArchitecture 
			        " > OS Name: " + $os.Caption 
			        " > OS Service Pack: " + $os.ServicePackMajorVersion 
			        " > CPU Max Clock Speed: " + $processor.MaxClockSpeed 
			        " > CPU Cores: " + $processor.NumberOfCores 
			        " > CPU Logical Cores: " + $processor.NumberOfLogicalProcessors 
			        " > Processor Manufacture: " + $processor.Manufacturer 
			        " > Memory Capacity: " + $memory.capacity 
					
					if ($GetDCInfo -eq $true) {
						# Loop through the pre-2012 DC's and get relevant info
					   	if ($(Test-WSMan -ComputerName $DCName)) {
							$sessions = New-PSSession $DCName
						    $dcoutdir = "$dcrootdir\$($DCName)"
							New-Item $dcoutdir -ItemType directory | Out-Null
						    Invoke-Command -Session $sessions {Get-EventLog -logname Application -EntryType Error,Warning -Newest 250 | select -Property TimeGenerated,EntryType,EventID,Message} | Sort-Object TimeGenerated -Descending | Export-Csv $dcoutdir\ApplicationLog.csv -NoTypeInformation
						    Invoke-Command -Session $sessions {Get-EventLog -logname System -EntryType Error,Warning -Newest 250 | select -Property TimeGenerated,EntryType,EventID,Message} | Sort-Object TimeGenerated -Descending | Export-Csv $dcoutdir\SystemLog.csv -NoTypeInformation
						    Invoke-Command -Session $sessions {Get-EventLog -logname "DFS Replication" -EntryType Error,Warning -Newest 250 | select -Property TimeGenerated,EntryType,EventID,Message} | Sort-Object TimeGenerated -Descending | Export-Csv $dcoutdir\DFSReplicationLog.csv -NoTypeInformation
						    Invoke-Command -Session $sessions {Get-EventLog -logname "Directory Service" -EntryType Error,Warning -Newest 250 | select -Property TimeGenerated,EntryType,EventID,Message} | Sort-Object TimeGenerated -Descending | Export-Csv $dcoutdir\DirectoryServiceLog.csv -NoTypeInformation
						    Invoke-Command -Session $sessions {systeminfo} | Out-File $dcoutdir\systeminfo.txt  
						    Invoke-Command -Session $sessions {dcdiag} | Out-File $dcoutdir\dcdiag.txt  
						    Invoke-Command -Session $sessions {netsh interface ipv4 show config} | Out-File $dcoutdir\NetInterface.txt 
							$DNSValue = Invoke-Command -ScriptBlock{$DNSState = Get-WindowsFeature -NAME DNS;return $DNSState} -Session $sessions 
							if ($DNSValue){
								Invoke-Command -Session $sessions {dnscmd /enumzones} | Out-File $dcoutdir\dnszones.txt
							    Invoke-Command -Session $sessions {dnscmd /info} | Out-File $dcoutdir\dnsinfo.txt
								Invoke-Command -Session $sessions {Get-EventLog -logname "DNS Server" -EntryType Error,Warning -Newest 250 | select -Property TimeGenerated,EntryType,EventID,Message} | Sort-Object TimeGenerated -Descending | Export-Csv $dcoutdir\DNSServerLog.csv -NoTypeInformation
							}				    
							Invoke-Command -Session $sessions {w32tm /dumpreg /subkey:parameters} | Out-File $dcoutdir\timeconfig.txt
						    Invoke-Command -Session $sessions {repadmin /replsummary} | Out-File $dcoutdir\replsummary.txt
						    Invoke-Command -Session $sessions {repadmin /showreps} | Out-File $dcoutdir\showreps.txt
						    Invoke-Command -Session $sessions {net statistics server} | Out-File $dcoutdir\netstatistics.txt 
							Invoke-Command -Session $sessions {Get-ADFineGrainedPasswordPolicy -Filter * | select -Property Name,ComplexityEnabled,LockoutDuration,LockoutObservationWindow,LockoutThreshold,MaxPasswordAge,MinPasswordAge,MinPasswordLength,PasswordHistoryCount,Precedence} | Out-File $dcoutdir\ADFineGrainPassPolicy.txt 
						    Invoke-Command –session $sessions {Get-Content “c:\windows\debug\netlogon.log”} | Out-File $dcoutdir\NoClientSite.txt
							Remove-PSSession -ComputerName $DCName
						}
						else {
							$dcoutdir = "$dcrootdir\$($DCName)"
							New-Item $dcoutdir -ItemType directory | Out-Null
							Write-Host ""
							Write-Host "Error: WinRM cannot connect to $DCName. This could be due to Windows Remote Management (WS-Management) service not being started, due Windows Firewall rules, etc.  The following information will not be collected: Event Log Information, systeminfo, dcdiag, netsh, dnscmd, w32tm, repadmin, netstatistics, and AD Fine Grained Password Policy." -ForegroundColor Red
							Write-Host ""
							$ErrorMsg = "WinRM cannot connect to $DCName. This could be due to Windows Remote Management (WS-Management) service not being started, due Windows Firewall rules, etc.  The following information will not be collected: Event Log Information, systeminfo, dcdiag, netsh, dnscmd, w32tm, repadmin, netstatistics, and AD Fine Grained Password Policy."
							$ErrorMsg | Out-File $dcoutdir\WinRMFailure.txt
						}
					}
				}
				elseif ($DCPS.operatingsystem -match '2012') {
					$Opt = New-CimSessionOption -Protocol Dcom
					$Session = New-CimSession -ComputerName $DCName -SessionOption $Opt
					$os = Get-CimInstance -Class Win32_OperatingSystem –CimSession $session
					$os1 = Get-CimInstance -class Win32_PhysicalMemory –CimSession $session |Measure-Object -Property capacity -Sum
					$Bios = Get-CimInstance -Class Win32_BIOS –CimSession $session
					$processor = Get-CimInstance -Class Win32_Processor –CimSession $session

					$TotalAvailMemory = ([math]::round(($OS1.Sum / 1GB),2))
					" > OS Architecture: " + $os.OsArchitecture 
			        " > OS Name: " + $os.Name.split("|")[0]
			        " > OS Service Pack: " + $os.ServicePackMajorVersion 
			        " > CPU Max Clock Speed: " + $processor.MaxClockSpeed 
			        " > CPU Cores: " + $processor.NumberOfCores 
			        " > CPU Logical Cores: " + $processor.NumberOfLogicalProcessors 
			        " > Processor Manufacture: " + $processor.Manufacturer 
			        " > Memory Capacity: " + ([math]::round(($os1.Sum / 1GB),2))

					if ($GetDCInfo -eq $true) {
						# Loop through the pre-2012 DC's and get relevant info
					   	if ($(Test-WSMan -ComputerName $DCName)) {
							$sessions = New-PSSession $DCName
						    $dcoutdir = "$dcrootdir\$($DCName)"
							New-Item $dcoutdir -ItemType directory | Out-Null
		 					Invoke-Command -Session $sessions {Get-WindowsFeature | where {$_.InstallState -eq "Installed"} | select -Property DisplayName,FeatureType} | sort FeatureType,DisplayName -Descending | Export-Csv $dcoutdir\DCWindowsFeatures.csv -NoTypeInformation
							Invoke-Command -Session $sessions {Get-EventLog -logname Application -EntryType Error,Warning -Newest 250 | select -Property TimeGenerated,EntryType,EventID,Message} | Sort-Object TimeGenerated -Descending | Export-Csv $dcoutdir\ApplicationLog.csv -NoTypeInformation
						    Invoke-Command -Session $sessions {Get-EventLog -logname System -EntryType Error,Warning -Newest 250 | select -Property TimeGenerated,EntryType,EventID,Message} | Sort-Object TimeGenerated -Descending | Export-Csv $dcoutdir\SystemLog.csv -NoTypeInformation
						    Invoke-Command -Session $sessions {Get-EventLog -logname "DFS Replication" -EntryType Error,Warning -Newest 250 | select -Property TimeGenerated,EntryType,EventID,Message} | Sort-Object TimeGenerated -Descending | Export-Csv $dcoutdir\DFSReplicationLog.csv -NoTypeInformation
						    Invoke-Command -Session $sessions {Get-EventLog -logname "Directory Service" -EntryType Error,Warning -Newest 250 | select -Property TimeGenerated,EntryType,EventID,Message} | Sort-Object TimeGenerated -Descending | Export-Csv $dcoutdir\DirectoryServiceLog.csv -NoTypeInformation
						    Invoke-Command -Session $sessions {systeminfo} | Out-File $dcoutdir\systeminfo.txt  
						    Invoke-Command -Session $sessions {dcdiag} | Out-File $dcoutdir\dcdiag.txt  
						    Invoke-Command -Session $sessions {netsh interface ipv4 show config} | Out-File $dcoutdir\NetInterface.txt 
							$DNSValue = Invoke-Command -ScriptBlock{$DNSState = Get-WindowsFeature -NAME DNS;return $DNSState} -Session $sessions 
							if ($DNSValue){
								Invoke-Command -Session $sessions {dnscmd /enumzones} | Out-File $dcoutdir\dnszones.txt
							    Invoke-Command -Session $sessions {dnscmd /info} | Out-File $dcoutdir\dnsinfo.txt
								Invoke-Command -Session $sessions {Get-EventLog -logname "DNS Server" -EntryType Error,Warning -Newest 250 | select -Property TimeGenerated,EntryType,EventID,Message} | Sort-Object TimeGenerated -Descending | Export-Csv $dcoutdir\DNSServerLog.csv -NoTypeInformation
							}	
						    Invoke-Command -Session $sessions {w32tm /dumpreg /subkey:parameters} | Out-File $dcoutdir\timeconfig.txt
						    Invoke-Command -Session $sessions {repadmin /replsummary} | Out-File $dcoutdir\replsummary.txt
						    Invoke-Command -Session $sessions {repadmin /showreps} | Out-File $dcoutdir\showreps.txt
						    Invoke-Command -Session $sessions {net statistics server} | Out-File $dcoutdir\netstatistics.txt 
						    Invoke-Command -Session $sessions {Get-ADTrust -filter * | select -Property Direction,ForestTransitive,Name,Source,Target} | Out-File $dcoutdir\ADTrusts.txt 					
						    Invoke-Command -Session $sessions {Get-ADReplicationSiteLink -Filter * | select -Property Name,Cost,ReplicationFrequencyInMinutes } | Out-File $dcoutdir\ADReplicationSiteLink.txt 					
							Invoke-Command -Session $sessions {Get-ADFineGrainedPasswordPolicy -Filter * | select -Property Name,ComplexityEnabled,LockoutDuration,LockoutObservationWindow,LockoutThreshold,MaxPasswordAge,MinPasswordAge,MinPasswordLength,PasswordHistoryCount,Precedence} | Out-File $dcoutdir\ADFineGrainPassPolicy.txt 
							Invoke-Command –session $sessions {Get-Content “c:\windows\debug\netlogon.log”} | Out-File $dcoutdir\NoClientSite.txt
							Remove-PSSession -ComputerName $DCName
						}
						else {
							$dcoutdir = "$dcrootdir\$($DCName)"
							New-Item $dcoutdir -ItemType directory | Out-Null
							Write-Host "Error: WinRM cannot connect to $DCName. This could be due to Windows Remote Management (WS-Management) service not being started, due Windows Firewall rules, etc.  The following information will not be collected: Event Log Information, systeminfo, dcdiag, netsh, dnscmd, w32tm, repadmin, netstatistics, AD Trusts, AD Site Links, and AD Fine Grained Password Policy." -ForegroundColor Red
							$ErrorMsg = "WinRM cannot connect to $DCName. This could be due to Windows Remote Management (WS-Management) service not being started, due Windows Firewall rules, etc.  The following information will not be collected: Event Log Information, systeminfo, dcdiag, netsh, dnscmd, w32tm, repadmin, netstatistics, AD Trusts, AD Site Links, and AD Fine Grained Password Policy."
							$ErrorMsg | Out-File $dcoutdir\WinRMFailure.txt
						}  
					}
				}

				if (($DC.OSVersion -like "*Windows Server 2003*") -or ($DC.OSVersion -like "*Windows 2000*")) {
					$ntps = "NTP Source Lookup Not Supported for Windows Server 2003 Operating Systems or Older."
				}
				else {
					$ntps = w32tm /query /computer:$DC /source
				}
				" > NTP Source: " + $ntps
			}
			else { 
				" > Ping: Unsuccessful" 
				" > OS Version: Server Unavailable. Unable to retrieve information."
				" > Service Pack: Server Unavailable. Unable to retrieve information."
				" > OS Architecture: Server Unavailable. Unable to retrieve information."
			}
			""
		}

		if ($GetComputerInfo -eq $true) {
			Write-Host " "
			Write-Host "- Starting to collect Computer Information from $DomainName" -ForegroundColor Green
			Write-Host " "
			# Find info about the computers in the domain
			$AllComputers = Get-ADComputer -filter * -Properties * -Server $DomainName
			$DomainNameOutput = "Domain: " + $DomainName
			$DomainNameOutput | Out-File $computeroutdir\ComputerCount.txt -Append
			"All Computers: $(($allcomputers).count)" | Out-File $computeroutdir\ComputerCount.txt -Append

			# All Computers
			foreach ($Computer in $AllComputers) {		
				$ActiveComputerDN = $Computer.DistinguishedName
				$Pos = $ActiveComputerDN.IndexOf(",")
				$OU = $ActiveComputerDN.substring($Pos+1)
				$ADobjComputers = New-Object PSObject
				$ADobjComputers | Add-Member NoteProperty "Domain" $DomainName
				$ADobjComputers | Add-Member NoteProperty "Computer" $Computer.Name
				$ADobjComputers | Add-Member NoteProperty "OperatingSystem" $Computer.OperatingSystem
				$ADobjComputers | Add-Member NoteProperty "OU" $OU
				$ADobjComputers | Add-Member NoteProperty "Enabled" $Computer.Enabled
				$ADobjComputers | Add-Member NoteProperty "LastLogonDate" $Computer.LastLogonDate
				$AllComputersList += $ADobjComputers
			}
			
			# Active Computers
			$ActiveWorkstations = $AllComputers | where {$_.OperatingSystem -notlike "*server*" -and $_.LastLogonDate -gt (get-date).AddMonths(-3) -and $_.Enabled -eq $true}
			"Active Workstations: $(($ActiveWorkstations).count)" | Out-File $computeroutdir\ComputerCount.txt -Append
			foreach ($ActiveComputer in $ActiveWorkstations) {	
				$ActiveWorkstationDN = $ActiveComputer.DistinguishedName
				$Pos = $ActiveWorkstationDN.IndexOf(",")
				$OU = $ActiveWorkstationDN.substring($Pos+1)
				$ADobjActiveComputers = New-Object PSObject
				$ADobjActiveComputers | Add-Member NoteProperty "Domain" $DomainName
				$ADobjActiveComputers | Add-Member NoteProperty "Computer" $ActiveComputer.Name
				$ADobjActiveComputers | Add-Member NoteProperty "OperatingSystem" $ActiveComputer.OperatingSystem
				$ADobjActiveComputers | Add-Member NoteProperty "OU" $OU
				$ADobjActiveComputers | Add-Member NoteProperty "Enabled" $ActiveComputer.Enabled
				$ADobjActiveComputers | Add-Member NoteProperty "LastLogonDate" $ActiveComputer.LastLogonDate
				$ActiveComputersList += $ADobjActiveComputers
			}

			# Active Servers
			$ActiveServers = $AllComputers | where {$_.OperatingSystem -like "*server*" -and $_.LastLogonDate -gt (get-date).AddMonths(-3) -and $_.Enabled -eq $true}
			"Active Servers: $(($ActiveServers).count)" | Out-File $computeroutdir\ComputerCount.txt -Append
			foreach ($ActiveServer in $ActiveServers) {
				$ActiveServerDN = $ActiveServer.DistinguishedName
				$Pos = $ActiveServerDN.IndexOf(",")
				$OU = $ActiveServerDN.substring($Pos+1)
				$ADobjActiveServers = New-Object PSObject
				$ADobjActiveServers | Add-Member NoteProperty "Domain" $DomainName
				$ADobjActiveServers | Add-Member NoteProperty "Computer" $ActiveComputer.Name
				$ADobjActiveServers | Add-Member NoteProperty "OperatingSystem" $ActiveServer.OperatingSystem
				$ADobjActiveServers | Add-Member NoteProperty "OU" $OU
				$ADobjActiveServers | Add-Member NoteProperty "Enabled" $ActiveServer.Enabled
				$ADobjActiveServers | Add-Member NoteProperty "LastLogonDate" $ActiveServer.LastLogonDate
				$ActiveServersList += $ADobjActiveServers
			}
			
			# Inactive Computers
			$InactiveComputers = Search-ADAccount -AccountInactive -ComputersOnly -TimeSpan 90.00:00:00 -Server $DomainName
			"Inactive Computers: $(($InactiveComputers).count)" | Out-File $computeroutdir\ComputerCount.txt -Append
			foreach ($InactiveComputer in $InactiveComputers) {	
				$InactiveComputerDN = $InactiveComputer.DistinguishedName
				$Pos = $InactiveComputerDN.IndexOf(",")
				$OU = $InactiveComputerDN.substring($Pos+1)
				$ADobjInactiveComputers = New-Object PSObject
				$ADobjInactiveComputers | Add-Member NoteProperty "Domain" $DomainName
				$ADobjInactiveComputers | Add-Member NoteProperty "Computer" $InactiveComputer.Name
				$ADobjInactiveComputers | Add-Member NoteProperty "OperatingSystem" $InactiveComputer.OperatingSystem
				$ADobjInactiveComputers | Add-Member NoteProperty "OU" $OU
				$ADobjInactiveComputers | Add-Member NoteProperty "Enabled" $InactiveComputer.Enabled
				$ADobjInactiveComputers | Add-Member NoteProperty "LastLogonDate" $InactiveComputer.LastLogonDate
				$InactiveComputerList += $ADobjInactiveComputers
			}
			
			# Disabled Computers
			$DisabledComputers = Search-ADAccount -AccountDisabled -ComputersOnly -Server $DomainName
			"Disabled Computers: $(($DisabledComputers).count)" | Out-File $computeroutdir\ComputerCount.txt -Append
			$AddSpaceToFile = " "
			$AddSpaceToFile | Out-File $computeroutdir\ComputerCount.txt -Append
			foreach ($DisabledComputer in $DisabledComputers) {
				$DisabledComputerDN = $DisabledComputer.DistinguishedName
				$Pos = $DisabledComputerDN.IndexOf(",")
				$OU = $DisabledComputerDN.substring($Pos+1)
				$ADobjDisabledComputers = New-Object PSObject
				$ADobjDisabledComputers | Add-Member NoteProperty "Domain" $DomainName
				$ADobjDisabledComputers | Add-Member NoteProperty "Computer" $DisabledComputer.Name
				$ADobjDisabledComputers | Add-Member NoteProperty "OperatingSystem" $DisabledComputer.OperatingSystem
				$ADobjDisabledComputers | Add-Member NoteProperty "OU" $OU
				$ADobjDisabledComputers | Add-Member NoteProperty "Enabled" $DisabledComputer.Enabled
				$ADobjDisabledComputers | Add-Member NoteProperty "LastLogonDate" $DisabledComputer.LastLogonDate
				$DisabledComputerList += $ADobjDisabledComputers
			}
		}
		if ($GetUserInfo -eq $true) {
			Write-Host " "
			Write-Host "- Starting to collect User Information from $DomainName" -ForegroundColor Green
			Write-Host " "
			$ADUsers = Get-ADUser -Server $DomainName -Filter *
			
			# Gets all users and specific attributes				
			foreach ($ADUser in $ADUsers) {
				$dn = $ADUser.DistinguishedName
				$user = [ADSI]"LDAP://$dn"
				$acl = $user.objectSecurity
				# the following indicates whether inherited rights checkbox is cleared
				$inherited = $acl.AreAccessRulesProtected
				$inheritedstatus = if ($inherited -eq $true) { "Not Inheriting" } elseif ($inherited -eq $false) { "Is Inheriting" }
				# the following indicates whether AdminCount is set
				$admincount = $user.admincount
				$admincountstatus = if ($user.admincount -ne $null) { "1" } else { "<Not Set>" }
				$DisplayName = $user.DisplayName.ToString()
				$GivenName = $user.GivenName.ToString()
				$Surname = $user.Surname.ToString()
				$SamAccountName = $user.SamAccountName.ToString()
				$Enabled = $user.Enabled.ToString()
				$AccountExpirationDate = $user.AccountExpirationDate.ToString()
				$deliverAndRedirect = if ($User.deliverAndRedirect -like "*True*") { "TRUE" } elseif ($user.deliverAndRedirect -like "*False*") {"FALSE"} else { "<Not Set>" }
				$altRecipient = if ($user.altRecipient) { $user.altRecipient } else { "<Not Set>" }
				$UPN = $user.userPrincipalName.toString()
				$targetAddress = if ($user.targetAddress) { $user.targetAddress.toString() } else { "<Not Set>" }
				$mail = $user.mail.toString()
				
				$Pos = $dn.IndexOf(",")
				$OU = $dn.substring($Pos+1)
				
				$ADUserobj = New-Object PSObject
				$ADUserobj | Add-Member NoteProperty "Domain" $DomainName
				$ADUserobj | Add-Member NoteProperty "DisplayName" $DisplayName
				$ADUserobj | Add-Member NoteProperty "GivenName" $GivenName
				$ADUserobj | Add-Member NoteProperty "Surname" $Surname
				$ADUserobj | Add-Member NoteProperty "SamAccountName" $SamAccountName
				$ADUserobj | Add-Member NoteProperty "OU" $OU
				$ADUserobj | Add-Member NoteProperty "UserPrincipalName" $UPN
				$ADUserobj | Add-Member NoteProperty "Enabled" $Enabled
				$ADUserobj | Add-Member NoteProperty "AccountExpirationDate" $AccountExpirationDate
				$ADUserobj | Add-Member NoteProperty "mail" $mail	
				$ADUserobj | Add-Member NoteProperty "deliverAndRedirect" $deliverAndRedirect
				$ADUserobj | Add-Member NoteProperty "altRecipient" $altRecipient	
				$ADUserobj | Add-Member NoteProperty "TargetAddress" $targetAddress
				$ADUserobj | Add-Member NoteProperty "Inherited" $inheritedstatus
				$ADUserobj | Add-Member NoteProperty "AdminCount" $admincountstatus
				$UserSecurity += $ADUserobj
			} 
		}
	}
	
	if ($GetUserInfo -eq $true) { $UserSecurity | Export-Csv -Path $useroutdir\AllUsers.csv -NoTypeInformation }

	if ($GetComputerInfo -eq $true) {
		$AllComputersList | Export-Csv -Path $computeroutdir\AllComputers.csv -NoTypeInformation
		$ActiveComputersList | Export-Csv -Path $computeroutdir\ActiveComputers.csv -NoTypeInformation
		$ActiveServersList | Export-Csv -Path $computeroutdir\ActiveServers.csv -NoTypeInformation
		$InactiveComputerList | Export-Csv -Path $computeroutdir\InactiveComputers.csv -NoTypeInformation
		$DisabledComputerList | Export-Csv -Path $computeroutdir\DisabledComputers.csv -NoTypeInformation
	}
	$DomainAdminList | Export-Csv $useroutdir\DomainAdminList.csv -Append -NoTypeInformation 
	$EEAdminList | Export-Csv $useroutdir\EnterpriseAdminList.csv -Append -NoTypeInformation 
	$SAdminList | Export-Csv $useroutdir\SchemaAdminList.csv -Append -NoTypeInformation 
	$AdminList | Export-Csv $useroutdir\AdministratorsList.csv -Append -NoTypeInformation 
	$PasswordPolicies | Export-Csv $adoutdir\DefaultDomainPWPolicy.csv -Append -NoTypeInformation 
}
Get-ADDomainInformation | Out-File $adoutdir\DomainInformation.txt

# Get all Active Directory Sites
$configNCDN = (Get-ADRootDSE).ConfigurationNamingContext
$siteContainerDN = ("CN=Sites," + $configNCDN)
$sites = Get-ADObject -SearchBase $siteContainerDN -filter { objectClass -eq "site" } -properties "siteObjectBL", "location", "description" | select Name, Location, Description

# Loop through all sites and get info for each site (DC's and Subnets etc..)
foreach($site in $sites){
    # information to be shared for both DC's and Subnets in each site
    $siteName =  $site.name 
    $configNCDN = (Get-ADRootDSE).ConfigurationNamingContext
    $siteContainerDN = ("CN=Sites," + $configNCDN)
    $serverContainerDN = "CN=Servers,CN=" + $siteName + "," + $siteContainerDN

    # Handle the DC's in the site first
    $dcsInSite = Get-ADObject -SearchBase $serverContainerDN -SearchScope OneLevel -filter { objectClass -eq "Server" } -Properties "DNSHostName" | Select Name

    # Create an object and format it to be dumped to CSV
    $export = @()
    $i = 0
    foreach($dc in $dcsInSite){
        $obj = New-Object System.Object
        # On the first line, show the site name
        if($i -eq 0){
            $obj | Add-Member -MemberType NoteProperty -Name Site -Value $siteName
            $obj | Add-Member -MemberType NoteProperty -Name DC -Value $dc.Name
            $i++
        }else{
            $obj | Add-Member -MemberType NoteProperty -Name DC -Value $dc.Name
        }
        $export += $obj
    }
    $export | Export-Csv $adoutdir\ADSites.csv -NoTypeInformation -append

    # Process the subnets now
    $siteDN = "CN=" + $siteName + "," + $siteContainerDN
    $siteObj = Get-ADObject -Identity $siteDN -properties "siteObjectBL", "description", "location" 
    foreach ($subnetDN in $siteObj.siteObjectBL) {
        $subnetsInSite = Get-ADObject -Identity $subnetDN -properties "siteObject", "description", "location" 
        # Create an object and format it to be dumped to CSV
        $export = @()
        foreach($subnet in $subnetsInSite){
            $obj = New-Object System.Object
            $obj | Add-Member -MemberType NoteProperty -Name Subnet -Value $subnet.Name
            $obj | Add-Member -MemberType NoteProperty -Name Site -Value $siteName
            $export += $obj
        }
        $export | Export-Csv $adoutdir\ADSubnets.csv -NoTypeInformation -append
    }
    
}

# get info about the DC's
$DCs = (Get-ADForest).Domains | %{ Get-ADDomainController –Filter * -Server $_ }  
$DCs | select -Property Name,Domain,Forest,IPv4Address,IsGlobalCatalog,OperatingSystem,OperatingSystemVersion,Site | Export-Csv $adoutdir\DomainControllers.csv -NoTypeInformation

Write-Host ""
Write-Host "- Completed!" -NoNewline "Ouput from the script is located in $filepath." -ForegroundColor Green
Write-Host ""