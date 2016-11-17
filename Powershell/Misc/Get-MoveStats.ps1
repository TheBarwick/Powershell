#########################################################################################
# COMPANY: CDW										                                    #
# NAME: Get-MoveStats.ps1			         			                                #
# 											                                            #
# AUTHOR:  Dean Sesko									                                #
# 											                                            #
# DATE:  08/28/2014									                                    #
# EMAIL: Dean.Sesko@S3.cdw.com							                                #
# 											                                            #
# COMMENT:  Get Mailbox Move Stats                                               	    #
#											                                            #
# VERSION HISTORY									                                    #
# 1.0 07/30/2015 Initial Version.							                            #
#											                                            #
#########################################################################################
param ([Parameter(Mandatory = $False)]
	[string]$MenuChoice
)
#	Setup Shell
$bg = (Get-Host).UI.RawUI
$BG.BackgroundColor = "black"

if ($MenuChoice -eq ""){
    $exitSwitch = $false
}
else {
 $exitSwitch = $true
}


# Check for Office 365 Connection
Try { $test = Get-MsolDomain -ErrorAction stop | where { ($_.name -like "*.mail.onmicrosoft.com") } }
Catch { Invoke-Expression .\Connect365.ps1 }
Finally { }

Function Show-Menu {
	cls
	Write-Host  $("-" * 85) -ForegroundColor Green
	Write-Host ""
	Write-Host "                 Office 365 Mailbox Move Statistics" -ForegroundColor Green
	Write-Host ""
	Write-Host  $("-" * 85) -ForegroundColor Green
	Write-Host ""
	Write-Host "1: Get All Mailbox Move statistics"
	Write-Host ""
	Write-Host "2: Get All In Progress Mailbox Move statistics"
	Write-Host ""
	Write-Host "3: Get All But Complted Mailbox Move Statistics"
	Write-Host ""
	Write-Host  $("-" * 85) -ForegroundColor Green
	Write-Host ""
	Write-Host ""
	$menudata += "Please Enter 1-3 or Q to Quit"
	$MyMenu = $menudata
	$MenuChoice =Read-Host -Prompt $MyMenu -vb
return $MenuChoice
	
	#End Function
}
function MoveStatsinfo ($moveStats) {
	
	
	if ($moveStats) {
		$moveStats | Get-MoveRequestStatistics | Sort-Object PercentComplete | ft DisplayName, PercentComplete, StatusDetail, By* -AutoSize
		Write-Host ""
		Write-Host ""
		Write-Host ""
		Write-Host $moveStats.count " Move Request Have Been Submited...." -ForegroundColor Green -BackgroundColor Black
		Write-Host ""
	}
	
	Else {
		Write-Host ""
		Write-Host ""
		Write-Host ""
		Write-Host "0 Move Request Have Been Submited...." -ForegroundColor Green -BackgroundColor Black
		Write-Host ""
	}
}

function ShowHeader {
	cls
	Write-host ""
	Write-Host  $("-" * 85) -ForegroundColor Green
	Write-host ""
	Write-Host "Getting MoveRequest.  This process can take some time"
	Write-host ""
	Write-Host  $("-" * 85) -ForegroundColor Green
	
	
}
cls



Do {
	
	if ($MenuChoice -eq "" -or $exitSwitch -eq $False){
	$MenuChoice = Show-Menu

}
	Switch ($MenuChoice) {
		"1" {
			ShowHeader
			$moves = Get-MoveRequest
			$exitSwitch = $true
			MoveStatsinfo $moves
					}
		"2" {
			ShowHeader
			$moves = Get-MoveRequest | Where { $_.status -eq "InProgress" -or $_.status -eq "CompletionInProgress" }
			$exitSwitch = $true
			MoveStatsinfo $moves
			
		}
		"3" {
			ShowHeader
			$moves = Get-MoveRequest | Where { $_.status -ne "Completed" }
			$exitSwitch = $true
			MoveStatsinfo $moves
			
		}
		
		"Q" {
			cls
			Write-Host ""
			Write-Host ""
			Write-Host "Good Bye"
			Write-Host ""
			$ExitScript = $True
			$exitSwitch = $true
		}
		default { $exitSwitch = $false }
		
	}
	


}

While ($exitSwitch -ne $true)
