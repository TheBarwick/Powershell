#########################################################################################
# COMPANY: CDW										                                    #
# NAME: Get-MigrationStatusReprot.pst   						                        #
# 											                                            #
# AUTHOR:  Dean Sesko									                                #
# 											                                            #
# DATE:  08/28/2014									                                    #
# EMAIL: Dean.SEsko@s3.CDW.com								                            #
# 											                                            #
# COMMENT:  Script to connect to Office 365 Administrator Shell and Generate a          # 
#           Migration HTML Report based on the current Get-Moverequest Status           #
#											                                            #
# VERSION HISTORY									                                    #
# 1.0 06/10/2014 Initial Version.							                            #
#											                                            #
#########################################################################################
$bg = (Get-Host).UI.RawUI
$BG.BackgroundColor = "black"
$HTMLFile = "O365_Migration_Status.htm"

#Test Office 365 Connection 
Try { $test = Get-MsolDomain -ErrorAction stop | where { ($_.name -like "*.mail.onmicrosoft.com") } }
Catch { Invoke-Expression .\Connect365.ps1 }
Finally { }

$TenantRoutingDomain = $TenantDom = Get-MsolDomain | where { $_.name -like "*.mail.onmicrosoft.com" }



#Setup HTML Header
$Output = "
<html> 
	<head> 
		<body>
		<font size=""5"" face=""Arial,sans-serif"">
		<h3 align=""Left"">Mailbox Move Report</h3>
		<h5 align=""Left"">Generated $((Get-Date).ToString())</h5>
		</font>
		<br>
		<br>
		<style type=text/css> 
			#menus { 
				overflow:hidden; 
				} 
			#left-table, #right-table { 
			float:left; 
				} 
			#left-table { 
			margin-right:1px; 
				} 
		h1 {font-size:30px;}
		h2 {font-size:25px;}
		</style>
		</head>
	<body>
<div id=""tables""> 
<table border=""0"" width=""100%"">
<tr>
<td>"
#Connecting to office 365

$moves = Get-MoveRequest


# First run to get Current mailbox moves in Progress
$Output += "<table border=""1"" id=""Left-table"" cellpadding=""3"" style=""font-size:12pt;font-family:Arial,sans-serif"">
<tr bgcolor=""#0000FF"">
<th colspan=""""><font color=""#ffffff"">In Progress:</font></th>
<tr bgcolor=""#00CC00""></tr>"
$Count = 0
foreach ($move in $moves) {
	
	if ($move.Status -eq "InProgress" -or $move.Status -eq "CompletionInProgress" -or $move.status -eq "Queued") {
		
		$count++
		$Name = $move.name
		$output += "<tr><td> $Name </td> </tr>"
		
	}
	
}
if ($count -eq 0) {
	$output += "<tr><td> Not Available</td> </tr>"
}
$Output += "</table>"



# Second run to get Completed mailbox moves 
$Output += "<table border=""1"" id=""Right-table"" cellpadding=""3"" style=""font-size:12pt;font-family:Arial,sans-serif"">
<tr bgcolor=""#009900"">
<th colspan=""""><font color=""#ffffff"">Completed:</font></th>
<tr bgcolor=""#00CC00""></tr>"
$Count = 0
foreach ($move in $moves) {
	
	
	if ($move.Status -eq "Completed" -or $move.Status -eq "CompletedWithWarning") {
		
		$count++
		$Name = $move.name
		$output += "<tr><td> $Name </td> </tr>"
		
	}
	
}
if ($count -eq 0) {
	$output += "<tr><td> Not Available</td> </tr>"
}
$Output += "</table>"



# Third run to get Susppended mailbox moves 
$Output += "<table border=""1"" id=""left-table"" cellpadding=""3"" style=""font-size:12pt;font-family:Arial,sans-serif"">
<tr bgcolor=""#FF0000"">
<th colspan=""""><font color=""#ffffff"">Queued:</font></th>
<tr bgcolor=""#00CC00""></tr>"
$Count = 0
foreach ($move in $moves) {
	
	
	if ($move.Status -eq "Autosuspended" -or $move.Status -eq "Suspended") {
		$count++
		$Name = $move.name
		$output += "<tr><td> $Name </td> </tr>"
		
	}
	
}
if ($count -eq 0) {
	$output += "<tr><td> Not Available</td> </tr>"
}
$Output += "</div></table></body>"
$Output | out-file $HTMLFile

