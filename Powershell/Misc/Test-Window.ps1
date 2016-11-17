#========================================================================
# Date: 8/28/2011 6:20 PM
# Author: Brad Stevens
#========================================================================
#----------------------------------------------
#region Application Functions
#----------------------------------------------

function OnApplicationLoad {


	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null



}

function OnApplicationExit {

}


#----------------------------------------------
# Generated Form Function
#----------------------------------------------
function Call-ANUC_pff {

	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$formMain = New-Object System.Windows.Forms.Form
	$btnSubmitAll = New-Object System.Windows.Forms.Button
	$btnLast = New-Object System.Windows.Forms.Button
	$btnNext = New-Object System.Windows.Forms.Button
	$btnPrev = New-Object System.Windows.Forms.Button
	$btnFirst = New-Object System.Windows.Forms.Button
	$btnImportCSV = New-Object System.Windows.Forms.Button
	$lvCSV = New-Object System.Windows.Forms.ListView
	$txtUPN = New-Object System.Windows.Forms.TextBox
	$txtsAM = New-Object System.Windows.Forms.TextBox
	$txtDN = New-Object System.Windows.Forms.TextBox
	$cboDepartment = New-Object System.Windows.Forms.ComboBox
	$labelUserPrincipalName = New-Object System.Windows.Forms.Label
	$labelSamAccountName = New-Object System.Windows.Forms.Label
	$labelDisplayName = New-Object System.Windows.Forms.Label
	$SB = New-Object System.Windows.Forms.StatusBar
	$cboSite = New-Object System.Windows.Forms.ComboBox
	$labelSite = New-Object System.Windows.Forms.Label
	$cboDescription = New-Object System.Windows.Forms.ComboBox
	$txtPassword = New-Object System.Windows.Forms.TextBox
	$labelPassword = New-Object System.Windows.Forms.Label
	$txtDomain = New-Object System.Windows.Forms.TextBox
	$labelCurrentDomain = New-Object System.Windows.Forms.Label
	$txtPostalCode = New-Object System.Windows.Forms.TextBox
	$txtState = New-Object System.Windows.Forms.TextBox
	$txtCity = New-Object System.Windows.Forms.TextBox
	$txtStreetAddress = New-Object System.Windows.Forms.TextBox
	$txtOffice = New-Object System.Windows.Forms.TextBox
	$txtCompany = New-Object System.Windows.Forms.TextBox
	$txtTitle = New-Object System.Windows.Forms.TextBox
	$txtOfficePhone = New-Object System.Windows.Forms.TextBox
	$txtLastName = New-Object System.Windows.Forms.TextBox
	$cboPath = New-Object System.Windows.Forms.ComboBox
	$labelOU = New-Object System.Windows.Forms.Label
	$txtFirstName = New-Object System.Windows.Forms.TextBox
	$labelPostalCode = New-Object System.Windows.Forms.Label
	$labelState = New-Object System.Windows.Forms.Label
	$labelCity = New-Object System.Windows.Forms.Label
	$labelStreetAddress = New-Object System.Windows.Forms.Label
	$labelOffice = New-Object System.Windows.Forms.Label
	$labelCompany = New-Object System.Windows.Forms.Label
	$labelDepartment = New-Object System.Windows.Forms.Label
	$labelTitle = New-Object System.Windows.Forms.Label
	$btnSubmit = New-Object System.Windows.Forms.Button
	$labelDescription = New-Object System.Windows.Forms.Label
	$labelOfficePhone = New-Object System.Windows.Forms.Label
	$labelLastName = New-Object System.Windows.Forms.Label
	$labelFirstName = New-Object System.Windows.Forms.Label
	$menustrip1 = New-Object System.Windows.Forms.MenuStrip
	$fileToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$formMode = New-Object System.Windows.Forms.ToolStripMenuItem
	$CSVTemplate = New-Object System.Windows.Forms.SaveFileDialog
	$OFDImportCSV = New-Object System.Windows.Forms.OpenFileDialog
	$CreateCSVTemplate = New-Object System.Windows.Forms.ToolStripMenuItem
	$MenuExit = New-Object System.Windows.Forms.ToolStripMenuItem
	$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
	
	
	$formMain_Load={
		
		$formMain.Text = $formMain.Text + " " + $XML.Options.Version
		
		Write-Verbose "Adding OUs to combo box"
	    $XML.Options.Domains.Domain | ?{$_.Name -match $txtDomain.Text} | Select -ExpandProperty Path | %{$cboPath.Items.Add($_)}
		
		Write-Verbose "Adding descriptions to combo box"
		$XML.Options.Descriptions.Description | %{$cboDescription.Items.Add($_)}
		
		Write-Verbose "Adding sites to combo box"
		$XML.Options.Locations.Location | %{$cboSite.Items.Add($_.Site)}
		
		Write-Verbose "Adding departments to combo box"
		$XML.Options.Departments.Department | %{$cboDepartment.Items.Add($_)}
		
		Write-Verbose "Setting default fields"
		$txtDomain.Text = $XML.Options.Default.Domain
	    $cboPath.SelectedItem = $XML.Options.Default.Path
		$txtFirstName.Text = $XML.Options.Default.FirstName
		$txtLastName.Text = $XML.Options.Default.LastName
		$txtOffice.Text = $XML.Options.Default.Office
		$txtTitle.Text = $XML.Options.Default.Title
		$cboDescription.SelectedItem = $XML.Options.Default.Description
		$cboDepartment.SelectedItem = $XML.Options.Default.Department
		$txtCompany.Text = $XML.Options.Default.Company
		$txtOfficePhone.Text = $XML.Options.Default.Phone
		$cboSite.SelectedItem = $XML.Options.Default.Site
		$txtStreetAddress.Text = $XML.Options.Default.StreetAddress
		$txtCity.Text = $XML.Options.Default.City
		$txtState.Text = $XML.Options.Default.State
		$txtPostalCode.Text = $XML.Options.Default.PostalCode
		$txtPassword.Text = $XML.Options.Default.Password
		
		Write-Verbose "Creating CSV Headers"
		$Headers = @('ID','Domain','Path','FirstName','LastName','Office','Title','Description','Department','Company','Phone','StreetAddress','City','State','PostalCode','Password','sAMAccountName','userPrincipalName','DisplayName')
		$Headers| %{[Void]$lvCSV.Columns.Add($_)}
	}
	
	$btnSubmit_Click={
		
		$Domain=$txtDomain.Text
		$Path=$cboPath.Text
		$GivenName = $txtFirstName.Text
		$Surname = $txtLastName.Text
		$OfficePhone = $txtOfficePhone.Text
		$Description = $cboDescription.Text
		$Title = $txtTitle.Text
		$Department = $cboDepartment.Text
		$Company = $txtCompany.Text
		$Office = $txtOffice.Text
		$StreetAddress = $txtStreetAddress.Text
		$City = $txtCity.Text
		$State = $txtState.Text
		$PostalCode = $txtPostalCode.Text
	
		if($XML.Options.Settings.Password.ChangeAtLogon -eq "True"){$ChangePasswordAtLogon = $True}
        else{$ChangePasswordAtLogon = $false}
		
        if($XML.Options.Settings.AccountStatus.Enabled -eq "True"){$Enabled = $True}
        else{$Enabled = $false}
	
		$Name="$GivenName $Surname"
		
        if($XML.Options.Settings.sAMAccountName.Generate -eq $True){$sAMAccountName = Set-sAMAccountName}
		else{$sAMAccountName = $txtsAM.Text}

        if($XML.Options.Settings.uPN.Generate -eq $True){$userPrincipalName = Set-UPN}
        else{$userPrincipalName = $txtuPN.Text}
		
        if($XML.Options.Settings.DisplayName.Generate -eq $True){$DisplayName = Set-DisplayName}
        else{$DisplayName = $txtDN.Text}

		$AccountPassword = $txtPassword.text | ConvertTo-SecureString -AsPlainText -Force
	
		$User = @{
		    Name = $Name
		    GivenName = $GivenName
		    Surname = $Surname
		    Path = $Path
		    samAccountName = $samAccountName
		    userPrincipalName = $userPrincipalName
		    DisplayName = $DisplayName
		    AccountPassword = $AccountPassword
		    ChangePasswordAtLogon = $ChangePasswordAtLogon
		    Enabled = $Enabled
		    OfficePhone = $OfficePhone
		    Description = $Description
		    Title = $Title
		    Department = $Department
		    Company = $Company
		    Office = $Office
		    StreetAddress = $StreetAddress
		    City = $City
		    State = $State
		    PostalCode = $PostalCode
		    }
		$SB.Text = "Creating new user $sAMAccountName"
        $ADError = $Null
		New-ADUser @User -ErrorVariable ADError
        if ($ADerror){$SB.Text = "[$sAMAccountName] $ADError"}
        else{$SB.Text = "$sAMAccountName created successfully."}
	}
	
	$txtDomain_SelectedIndexChanged={
		$cboPath.Items.Clear()
		Write-Verbose "Adding OUs to combo box"
	    $XML.Options.Domains.Domain | ?{$_.Name -match $txtDomain.Text} | Select -ExpandProperty Path | %{$cboPath.Items.Add($_)}	
		Write-Verbose "Creating required account fields"
		
        if ($XML.Options.Settings.DisplayName.Generate) {$txtDN.Text = Set-DisplayName}
        if ($XML.Options.Settings.sAMAccountName.Generate) {$txtsAM.Text = Set-sAMAccountName}
        if ($XML.Options.Settings.UPN.Generate) {$txtUPN.Text = Set-UPN}	
	}
	
	$cboSite_SelectedIndexChanged={
		Write-Verbose "Updating site fields with address information"
	    $Site = $XML.Options.Locations.Location | ?{$_.Site -match $cboSite.Text}
		$txtStreetAddress.Text = $Site.StreetAddress
		$txtCity.Text = $Site.City
		$txtState.Text = $Site.State
		$txtPostalCode.Text = $Site.PostalCode
	}
	
	$txtName_TextChanged={
		Write-Verbose "Creating required account fields"
        
        if ($XML.Options.Settings.DisplayName.Generate -eq $True) {$txtDN.Text = Set-DisplayName}
        if ($XML.Options.Settings.sAMAccountName.Generate -eq $True) {$txtsAM.Text = (Set-sAMAccountName)}
        if ($XML.Options.Settings.UPN.Generate -eq $True) {$txtUPN.Text = Set-UPN}
	}
	
	$createTemplateToolStripMenuItem_Click={
		$CSVTemplate.ShowDialog()
	}
	
	$CSVTemplate_FileOk=[System.ComponentModel.CancelEventHandler]{
		"" |
		Select Domain,Path,FirstName,LastName,Office,Title,Description,Department,Company,Phone,StreetAddress,City,State,PostalCode,Password,sAMAccountName,userPrincipalName,DisplayName |
		Export-CSV $CSVTemplate.FileName -NoTypeInformation	
	}

	$MenuExit_Click={
		$formMain.Close()
	}
	
	$btnSubmitAll_Click={
		$lvCSV.Items | %{
			
			$Domain = $_.Subitems[1].Text
			$Path = $_.Subitems[2].Text
			$GivenName = $_.Subitems[3].Text
			$Surname = $_.Subitems[4].Text
			$OfficePhone = $_.Subitems[5].Text
			$Title = $_.Subitems[6].Text
			$Description = $_.Subitems[7].Text
			$Department = $_.Subitems[8].Text
			$Company = $_.Subitems[9].Text
			$Office = $_.Subitems[10].Text
			$StreetAddress = $_.Subitems[11].Text
			$City = $_.Subitems[12].Text
			$State = $_.Subitems[13].Text
			$PostalCode = $_.Subitems[14].Text
	
			$Name = "$GivenName $Surname"

		    if($XML.Options.Settings.Password.ChangeAtLogon -eq "True"){$ChangePasswordAtLogon = $True}
            else{$ChangePasswordAtLogon = $false}
		
            if($XML.Options.Settings.AccountStatus.Enabled -eq "True"){$Enabled = $True}
            else{$Enabled = $false}
	
		    if($_.Subitems[16].Text -eq $null){$sAMAccountName = Set-sAMAccountName}
		    else{$sAMAccountName = $_.Subitems[16].Text}

            if($_.Subitems[17].Text -eq $null){$userPrincipalName = Set-UPN}
            else{$userPrincipalName = $_.Subitems[17].Text}
		
            if($_.Subitems[18].Text -eq $null){$DisplayName = Set-DisplayName}
            else{$DisplayName = $_.Subitems[18].Text}

			$AccountPassword = $_.Subitems[15].Text | ConvertTo-SecureString -AsPlainText -Force
	
			$User = @{
			    Name = $Name
			    GivenName = $GivenName
			    Surname = $Surname
			    Path = $Path
			    samAccountName = $samAccountName
			    userPrincipalName = $userPrincipalName
			    DisplayName = $DisplayName
			    AccountPassword = $AccountPassword
			    ChangePasswordAtLogon = $ChangePasswordAtLogon
			    Enabled = $Enabled
			    OfficePhone = $OfficePhone
			    Description = $Description
			    Title = $Title
			    Department = $Department
			    Company = $Company
			    Office = $Office
			    StreetAddress = $StreetAddress
			    City = $City
			    State = $State
			    PostalCode = $PostalCode
			    }
			$SB.Text = "Creating new user $sAMAccountName"
            $ADError = $Null
			New-ADUser @User -ErrorVariable ADError
            if ($ADerror){$SB.Text = "[$sAMAccountName] $ADError"}
            else{$SB.Text = "$sAMAccountName created successfully."}
		}
	}
	
	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$formMain.WindowState = $InitialFormWindowState
	}
	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$btnSubmitAll.remove_Click($btnSubmitAll_Click)
			$btnLast.remove_Click($btnLast_Click)
			$btnNext.remove_Click($btnNext_Click)
			$btnPrev.remove_Click($btnPrev_Click)
			$btnFirst.remove_Click($btnFirst_Click)
			$btnImportCSV.remove_Click($btnImportCSV_Click)
			$lvCSV.remove_SelectedIndexChanged($lvCSV_SelectedIndexChanged)
			$cboSite.remove_SelectedIndexChanged($cboSite_SelectedIndexChanged)
			$txtDomain.remove_SelectedIndexChanged($txtDomain_SelectedIndexChanged)
			$txtLastName.remove_TextChanged($txtName_TextChanged)
			$txtFirstName.remove_TextChanged($txtName_TextChanged)
			$btnSubmit.remove_Click($btnSubmit_Click)
			$formMain.remove_Load($formMain_Load)
			$formMode.remove_Click($formMode_Click)
			$CSVTemplate.remove_FileOk($CSVTemplate_FileOk)
			$CreateCSVTemplate.remove_Click($createTemplateToolStripMenuItem_Click)
			$MenuExit.remove_Click($MenuExit_Click)
			$formMain.remove_Load($Form_StateCorrection_Load)
			$formMain.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch [Exception]
		{ }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	#
	# formMain
	#
	$formMain.Controls.Add($btnSubmitAll)
	$formMain.Controls.Add($btnLast)
	$formMain.Controls.Add($btnNext)
	$formMain.Controls.Add($btnPrev)
	$formMain.Controls.Add($btnFirst)
	$formMain.Controls.Add($btnImportCSV)
	$formMain.Controls.Add($lvCSV)
	$formMain.Controls.Add($txtDN)
	$formMain.Controls.Add($cboDepartment)
	$formMain.Controls.Add($labelDisplayName)
	$formMain.Controls.Add($SB)
	$formMain.Controls.Add($cboSite)
	$formMain.Controls.Add($labelSite)
	$formMain.Controls.Add($cboDescription)
	$formMain.Controls.Add($txtPassword)
	$formMain.Controls.Add($labelPassword)
	$formMain.Controls.Add($txtDomain)
	$formMain.Controls.Add($labelCurrentDomain)
	$formMain.Controls.Add($txtPostalCode)
	$formMain.Controls.Add($txtState)
	$formMain.Controls.Add($txtCity)
	$formMain.Controls.Add($txtStreetAddress)
	$formMain.Controls.Add($txtOffice)
	$formMain.Controls.Add($txtCompany)
	$formMain.Controls.Add($txtTitle)
	$formMain.Controls.Add($txtOfficePhone)
	$formMain.Controls.Add($txtLastName)
	$formMain.Controls.Add($cboPath)
	$formMain.Controls.Add($labelOU)
	$formMain.Controls.Add($txtFirstName)
	$formMain.Controls.Add($labelPostalCode)
	$formMain.Controls.Add($labelState)
	$formMain.Controls.Add($labelCity)
	$formMain.Controls.Add($labelStreetAddress)
	$formMain.Controls.Add($labelOffice)
	$formMain.Controls.Add($labelCompany)
	$formMain.Controls.Add($labelDepartment)
	$formMain.Controls.Add($labelTitle)
	$formMain.Controls.Add($btnSubmit)
	$formMain.Controls.Add($labelDescription)
	$formMain.Controls.Add($labelOfficePhone)
	$formMain.Controls.Add($labelLastName)
	$formMain.Controls.Add($labelFirstName)
	$formMain.Controls.Add($menustrip1)
	$formMain.AcceptButton = $btnSubmit
	$formMain.ClientSize = '395, 630'
	$System_Windows_Forms_MenuStrip_1 = New-Object System.Windows.Forms.MenuStrip
	$System_Windows_Forms_MenuStrip_1.Location = '0, 0'
	$System_Windows_Forms_MenuStrip_1.Name = ""
	$System_Windows_Forms_MenuStrip_1.Size = '271, 24'
	$System_Windows_Forms_MenuStrip_1.TabIndex = 1
	$System_Windows_Forms_MenuStrip_1.Visible = $False
	$formMain.MainMenuStrip = $System_Windows_Forms_MenuStrip_1
	$formMain.Name = "formMain"
	$formMain.ShowIcon = $False
	$formMain.StartPosition = 'CenterScreen'
	$formMain.Text = "Brad's Ultimate New User Creation"
	$formMain.add_Load($formMain_Load)
	#
	# btnSubmitAll
	#
	$btnSubmitAll.Location = '503, 100'
	$btnSubmitAll.Name = "btnSubmitAll"
	$btnSubmitAll.Size = '75, 25'
	$btnSubmitAll.TabIndex = 59
	$btnSubmitAll.Text = "Submit All"
	$btnSubmitAll.UseVisualStyleBackColor = $True
	$btnSubmitAll.Visible = $False
	$btnSubmitAll.add_Click($btnSubmitAll_Click)
	#
	# btnLast
	#
	$btnLast.Location = '472, 0'
	$btnLast.Name = "btnLast"
	$btnLast.Size = '30, 25'
	$btnLast.TabIndex = 58
	$btnLast.Text = ">>"
	$btnLast.UseVisualStyleBackColor = $True
	$btnLast.Visible = $False
	$btnLast.add_Click($btnLast_Click)
	#
	# btnNext
	#
	$btnNext.Location = '441, 0'
	$btnNext.Name = "btnNext"
	$btnNext.Size = '30, 25'
	$btnNext.TabIndex = 57
	$btnNext.Text = ">"
	$btnNext.UseVisualStyleBackColor = $True
	$btnNext.Visible = $False
	$btnNext.add_Click($btnNext_Click)
	#
	# btnPrev
	#
	$btnPrev.Location = '410, 0'
	$btnPrev.Name = "btnPrev"
	$btnPrev.Size = '30, 25'
	$btnPrev.TabIndex = 56
	$btnPrev.Text = "<"
	$btnPrev.UseVisualStyleBackColor = $True
	$btnPrev.Visible = $False
	$btnPrev.add_Click($btnPrev_Click)
	#
	# btnFirst
	#
	$btnFirst.Location = '379, 0'
	$btnFirst.Name = "btnFirst"
	$btnFirst.Size = '30, 25'
	$btnFirst.TabIndex = 55
	$btnFirst.Text = "<<"
	$btnFirst.UseVisualStyleBackColor = $True
	$btnFirst.Visible = $False
	$btnFirst.add_Click($btnFirst_Click)
	#
	# btnImportCSV
	#
	$btnImportCSV.Location = '303, 0'
	$btnImportCSV.Name = "btnImportCSV"
	$btnImportCSV.Size = '75, 25'
	$btnImportCSV.TabIndex = 54
	$btnImportCSV.Text = "Import CSV"
	$btnImportCSV.UseVisualStyleBackColor = $True
	$btnImportCSV.Visible = $False
	$btnImportCSV.add_Click($btnImportCSV_Click)
	#
	# lvCSV
	#
	$lvCSV.FullRowSelect = $True
	$lvCSV.GridLines = $True
	$lvCSV.Location = '305, 35'
	$lvCSV.Name = "lvCSV"
	$lvCSV.Size = '1150, 535'
	$lvCSV.TabIndex = 53
	$lvCSV.UseCompatibleStateImageBehavior = $False
	$lvCSV.View = 'Details'
	$lvCSV.Visible = $False
	$lvCSV.add_SelectedIndexChanged($lvCSV_SelectedIndexChanged)
	#
	# txtDN
	#
	$txtDN.Anchor = 'Top, Left, Right'
	$txtDN.Location = '118, 455'
	$txtDN.Name = "txtDN"
	$txtDN.Size = '250, 20'
	$txtDN.TabIndex = 49
	#
	# cboDepartment
	#
	$cboDepartment.Anchor = 'Top, Left, Right'
	$cboDepartment.FormattingEnabled = $True
	$cboDepartment.Location = '118, 235'
	$cboDepartment.Name = "cboDepartment"
	$cboDepartment.Size = '250, 21'
	$cboDepartment.TabIndex = 8
	#
	# labelDisplayName
	#
	$labelDisplayName.Location = '10, 455'
	$labelDisplayName.Name = "labelDisplayName"
	$labelDisplayName.Size = '250, 23'
	$labelDisplayName.TabIndex = 46
	$labelDisplayName.Text = "Display Name"
	$labelDisplayName.TextAlign = 'MiddleLeft'
	#
	# SB
	#
	$SB.Location = '0, 575'
	$SB.Name = "SB"
	$SB.Size = '304, 22'
	$SB.TabIndex = 45
	$SB.Text = "Ready"
	#
	# cboSite
	#
	$cboSite.Anchor = 'Top, Left, Right'
	$cboSite.FormattingEnabled = $True
	$cboSite.Location = '118, 320'
	$cboSite.Name = "cboSite"
	$cboSite.Size = '250, 21'
	$cboSite.TabIndex = 11
	$cboSite.add_SelectedIndexChanged($cboSite_SelectedIndexChanged)
	#
	# labelSite
	#
	$labelSite.Location = '10, 320'
	$labelSite.Name = "labelSite"
	$labelSite.Size = '250, 23'
	$labelSite.TabIndex = 44
	$labelSite.Text = "Site"
	$labelSite.TextAlign = 'MiddleLeft'
	#
	# cboDescription
	#
	$cboDescription.Anchor = 'Top, Left, Right'
	$cboDescription.FormattingEnabled = $True
	$cboDescription.Location = '118, 210'
	$cboDescription.Name = "cboDescription"
	$cboDescription.Size = '250, 21'
	$cboDescription.TabIndex = 7
	#
	# txtPassword
	#
	$txtPassword.Anchor = 'Top, Left, Right'
	$txtPassword.Location = '118, 547'
	$txtPassword.Name = "txtPassword"
	$txtPassword.Size = '250, 20'
	$txtPassword.TabIndex = 16
	$txtPassword.UseSystemPasswordChar = $True
	#
	# labelPassword
	#
	$labelPassword.Location = '10, 545'
	$labelPassword.Name = "labelPassword"
	$labelPassword.Size = '250, 23'
	$labelPassword.TabIndex = 41
	$labelPassword.Text = "Password"
	$labelPassword.TextAlign = 'MiddleLeft'
	#
	# txtDomain
	#
	$txtDomain.Anchor = 'Top, Left, Right'
	$txtDomain.Location = '118, 35'
	$txtDomain.Name = "txtDomain"
	$txtDomain.Size = '250, 20'
	$txtDomain.TabIndex = 14
	#
	# labelCurrentDomain
	#
	$labelCurrentDomain.Location = '10, 35'
	$labelCurrentDomain.Name = "labelCurrentDomain"
	$labelCurrentDomain.Size = '250, 23'
	$labelCurrentDomain.TabIndex = 39
	$labelCurrentDomain.Text = "Domain"
	$labelCurrentDomain.TextAlign = 'MiddleLeft'
	#
	# txtPostalCode
	#
	$txtPostalCode.Anchor = 'Top, Left, Right'
	$txtPostalCode.Location = '118, 420'
	$txtPostalCode.Name = "txtPostalCode"
	$txtPostalCode.Size = '250, 20'
	$txtPostalCode.TabIndex = 15
	#
	# txtState
	#
	$txtState.Anchor = 'Top, Left, Right'
	$txtState.Location = '118, 395'
	$txtState.Name = "txtState"
	$txtState.Size = '250, 20'
	$txtState.TabIndex = 14
	#
	# txtCity
	#
	$txtCity.Anchor = 'Top, Left, Right'
	$txtCity.Location = '118, 370'
	$txtCity.Name = "txtCity"
	$txtCity.Size = '250, 20'
	$txtCity.TabIndex = 13
	#
	# txtStreetAddress
	#
	$txtStreetAddress.Anchor = 'Top, Left, Right'
	$txtStreetAddress.Location = '118, 345'
	$txtStreetAddress.Name = "txtStreetAddress"
	$txtStreetAddress.Size = '250, 20'
	$txtStreetAddress.TabIndex = 12
	#
	# txtOffice
	#
	$txtOffice.Anchor = 'Top, Left, Right'
	$txtOffice.Location = '118, 160'
	$txtOffice.Name = "txtOffice"
	$txtOffice.Size = '250, 20'
	$txtOffice.TabIndex = 5
	#
	# txtCompany
	#
	$txtCompany.Anchor = 'Top, Left, Right'
	$txtCompany.Location = '118, 260'
	$txtCompany.Name = "txtCompany"
	$txtCompany.Size = '250, 20'
	$txtCompany.TabIndex = 9
	#
	# txtTitle
	#
	$txtTitle.Anchor = 'Top, Left, Right'
	$txtTitle.Location = '118, 185'
	$txtTitle.Name = "txtTitle"
	$txtTitle.Size = '250, 20'
	$txtTitle.TabIndex = 6
	#
	# txtOfficePhone
	#
	$txtOfficePhone.Anchor = 'Top, Left, Right'
	$txtOfficePhone.Location = '118, 285'
	$txtOfficePhone.Name = "txtOfficePhone"
	$txtOfficePhone.Size = '250, 20'
	$txtOfficePhone.TabIndex = 10
	#
	# txtLastName
	#
	$txtLastName.Anchor = 'Top, Left, Right'
	$txtLastName.Location = '118, 135'
	$txtLastName.Name = "txtLastName"
	$txtLastName.Size = '250, 20'
	$txtLastName.TabIndex = 4
	$txtLastName.add_TextChanged($txtName_TextChanged)
	#
	# cboPath
	#
	$cboPath.Anchor = 'Top, Left, Right'
	$cboPath.FormattingEnabled = $True
	$cboPath.Location = '118, 65'
	$cboPath.Name = "cboPath"
	$cboPath.Size = '250, 21'
	$cboPath.TabIndex = 2
	#
	# labelOU
	#
	$labelOU.Location = '10, 65'
	$labelOU.Name = "labelOU"
	$labelOU.Size = '100, 23'
	$labelOU.TabIndex = 24
	$labelOU.Text = "OU"
	$labelOU.TextAlign = 'MiddleLeft'
	#
	# txtFirstName
	#
	$txtFirstName.Anchor = 'Top, Left, Right'
	$txtFirstName.Location = '118, 110'
	$txtFirstName.Name = "txtFirstName"
	$txtFirstName.Size = '250, 20'
	$txtFirstName.TabIndex = 3
	$txtFirstName.add_TextChanged($txtName_TextChanged)
	#
	# labelPostalCode
	#
	$labelPostalCode.Location = '10, 420'
	$labelPostalCode.Name = "labelPostalCode"
	$labelPostalCode.Size = '100, 23'
	$labelPostalCode.TabIndex = 24
	$labelPostalCode.Text = "Postal Code"
	$labelPostalCode.TextAlign = 'MiddleLeft'
	#
	# labelState
	#
	$labelState.Location = '10, 395'
	$labelState.Name = "labelState"
	$labelState.Size = '100, 23'
	$labelState.TabIndex = 23
	$labelState.Text = "State"
	$labelState.TextAlign = 'MiddleLeft'
	#
	# labelCity
	#
	$labelCity.Location = '10, 370'
	$labelCity.Name = "labelCity"
	$labelCity.Size = '100, 23'
	$labelCity.TabIndex = 22
	$labelCity.Text = "City"
	$labelCity.TextAlign = 'MiddleLeft'
	#
	# labelStreetAddress
	#
	$labelStreetAddress.Location = '10, 345'
	$labelStreetAddress.Name = "labelStreetAddress"
	$labelStreetAddress.Size = '100, 23'
	$labelStreetAddress.TabIndex = 21
	$labelStreetAddress.Text = "Street Address"
	$labelStreetAddress.TextAlign = 'MiddleLeft'
	#
	# labelOffice
	#
	$labelOffice.Location = '10, 160'
	$labelOffice.Name = "labelOffice"
	$labelOffice.Size = '250, 23'
	$labelOffice.TabIndex = 20
	$labelOffice.Text = "Office"
	$labelOffice.TextAlign = 'MiddleLeft'
	#
	# labelCompany
	#
	$labelCompany.Location = '10, 260'
	$labelCompany.Name = "labelCompany"
	$labelCompany.Size = '100, 23'
	$labelCompany.TabIndex = 19
	$labelCompany.Text = "Company"
	$labelCompany.TextAlign = 'MiddleLeft'
	#
	# labelDepartment
	#
	$labelDepartment.Location = '10, 235'
	$labelDepartment.Name = "labelDepartment"
	$labelDepartment.Size = '100, 23'
	$labelDepartment.TabIndex = 18
	$labelDepartment.Text = "Department"
	$labelDepartment.TextAlign = 'MiddleLeft'
	#
	# labelTitle
	#
	$labelTitle.Location = '10, 185'
	$labelTitle.Name = "labelTitle"
	$labelTitle.Size = '100, 23'
	$labelTitle.TabIndex = 17
	$labelTitle.Text = "Title"
	$labelTitle.TextAlign = 'MiddleLeft'
	#
	# btnSubmit
	#
	$btnSubmit.Location = '118, 580'
	$btnSubmit.Name = "btnSubmit"
	$btnSubmit.Size = '250, 25'
	$btnSubmit.TabIndex = 17
	$btnSubmit.Text = "Create User"
	$btnSubmit.UseVisualStyleBackColor = $True
	$btnSubmit.add_Click($btnSubmit_Click)
	#
	# labelDescription
	#
	$labelDescription.Location = '10, 210'
	$labelDescription.Name = "labelDescription"
	$labelDescription.Size = '100, 23'
	$labelDescription.TabIndex = 15
	$labelDescription.Text = "Description"
	$labelDescription.TextAlign = 'MiddleLeft'
	#
	# labelOfficePhone
	#
	$labelOfficePhone.Location = '10, 285'
	$labelOfficePhone.Name = "labelOfficePhone"
	$labelOfficePhone.Size = '100, 23'
	$labelOfficePhone.TabIndex = 14
	$labelOfficePhone.Text = "Office Phone"
	$labelOfficePhone.TextAlign = 'MiddleLeft'
	#
	# labelLastName
	#
	$labelLastName.Location = '10, 135'
	$labelLastName.Name = "labelLastName"
	$labelLastName.Size = '100, 23'
	$labelLastName.TabIndex = 13
	$labelLastName.Text = "Last Name"
	$labelLastName.TextAlign = 'MiddleLeft'
	#
	# labelFirstName
	#
	$labelFirstName.Location = '10, 110'
	$labelFirstName.Name = "labelFirstName"
	$labelFirstName.Size = '100, 23'
	$labelFirstName.TabIndex = 12
	$labelFirstName.Text = "First Name"
	$labelFirstName.TextAlign = 'MiddleLeft'
	#
	# menustrip1
	#
	[void]$menustrip1.Items.Add($fileToolStripMenuItem)
	$menustrip1.Location = '0, 0'
	$menustrip1.Name = "menustrip1"
	$menustrip1.Size = '304, 24'
	$menustrip1.TabIndex = 52
	$menustrip1.Text = "menustrip1"
	#
	# fileToolStripMenuItem
	#
	[void]$fileToolStripMenuItem.DropDownItems.Add($formMode)
	[void]$fileToolStripMenuItem.DropDownItems.Add($CreateCSVTemplate)
	[void]$fileToolStripMenuItem.DropDownItems.Add($MenuExit)
	$fileToolStripMenuItem.Name = "fileToolStripMenuItem"
	$fileToolStripMenuItem.Size = '37, 20'
	$fileToolStripMenuItem.Text = "File"
	#
	# formMode
	#
	$formMode.Name = "formMode"
	$formMode.Size = '185, 22'
	$formMode.Text = "CSV Mode"
	$formMode.add_Click($formMode_Click)
	#
	# CSVTemplate
	#
	$CSVTemplate.CheckPathExists = $False
	$CSVTemplate.DefaultExt = "csv"
	$CSVTemplate.FileName = "ANUCusers.csv"
	$CSVTemplate.Filter = "CSV Files|*.csv|All Files|*.*"
	$CSVTemplate.ShowHelp = $True
	$CSVTemplate.Title = "Create CSV Template For ANUC"
	$CSVTemplate.add_FileOk($CSVTemplate_FileOk)
	#
	# OFDImportCSV
	#
	$OFDImportCSV.FileName = "C:\ANUC\AnucUsers.csv"
	$OFDImportCSV.ShowHelp = $True
	#
	# CreateCSVTemplate
	#
	$CreateCSVTemplate.Name = "CreateCSVTemplate"
	$CreateCSVTemplate.Size = '185, 22'
	$CreateCSVTemplate.Text = "Create CSV Template"
	$CreateCSVTemplate.add_Click($createTemplateToolStripMenuItem_Click)
	#
	# MenuExit
	#
	$MenuExit.Name = "MenuExit"
	$MenuExit.Size = '185, 22'
	$MenuExit.Text = "Exit"
	$MenuExit.add_Click($MenuExit_Click)
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $formMain.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$formMain.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$formMain.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	return $formMain.ShowDialog()

} #End Function

#Call OnApplicationLoad to initialize
if((OnApplicationLoad) -eq $true)
{
	#Call the form
	Call-ANUC_pff | Out-Null
	#Perform cleanup
	OnApplicationExit
}
