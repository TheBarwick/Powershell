function Call-BUAS_pff {

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
	$groupTools = New-Object System.Windows.Forms.GroupBox
	$btnRestart = New-Object System.Windows.Forms.Button
	$btnMsg = New-Object System.Windows.Forms.Button
	$btnCDrive = New-Object System.Windows.Forms.Button
	$btnRA = New-Object System.Windows.Forms.Button
	$btnRDP = New-Object System.Windows.Forms.Button
	$groupInfo = New-Object System.Windows.Forms.GroupBox
    $userInfo = New-Object System.Windows.Forms.GroupBox
	$btnServices = New-Object System.Windows.Forms.Button
	$btnProcesses = New-Object System.Windows.Forms.Button
	$btnStartupItems = New-Object System.Windows.Forms.Button
	$btnApplications = New-Object System.Windows.Forms.Button
	$btnLocalAdmins = New-Object System.Windows.Forms.Button
    $btnAddNewUser = New-Object System.Windows.Forms.Button
	$btnSystemInfo = New-Object System.Windows.Forms.Button
	$lvMain = New-Object System.Windows.Forms.ListView
	$btnSearch = New-Object System.Windows.Forms.Button
	$txtComputer = New-Object System.Windows.Forms.TextBox
	$SB = New-Object System.Windows.Forms.StatusBar
	$menu = New-Object System.Windows.Forms.MenuStrip
	$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuFileConnect = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuFileExit = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuView = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuViewEventVwr = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuViewServices = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuViewUser = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuViewWSUS = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuViewWSUSReport = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuViewWSUSUpdate = New-Object System.Windows.Forms.ToolStripMenuItem
	$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
	$cmsProcEnd = New-Object System.Windows.Forms.ToolStripMenuItem
	$cmsStartupRemove = New-Object System.Windows.Forms.ToolStripMenuItem
	$cmsAdminAdd = New-Object System.Windows.Forms.ToolStripMenuItem
	$cmsAdminRemove = New-Object System.Windows.Forms.ToolStripMenuItem
	$cmsAppUninstall = New-Object System.Windows.Forms.ToolStripMenuItem
	$cmsSelect = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuHelpAbout = New-Object System.Windows.Forms.ToolStripMenuItem
	$SBPStatus = New-Object System.Windows.Forms.StatusBarPanel
	$SBPBlog = New-Object System.Windows.Forms.StatusBarPanel
	$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------