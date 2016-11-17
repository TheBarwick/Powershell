function Select-FileDialog 
{
	param([string]$Title,[string]$Directory,[string]$Filter="CSV Files (*.csv)|*.csv")
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$objForm = New-Object System.Windows.Forms.OpenFileDialog
	$objForm.InitialDirectory = $Directory
	$objForm.Filter = $Filter
	$objForm.Title = $Title
	$objForm.ShowHelp = $true
	
	$Show = $objForm.ShowDialog()
	
	If ($Show -eq "OK")
	{
		Return $objForm.FileName
	}
	Else
	{
		Exit
	}
}

function CountDown($waitMinutes) {
    
	$startTime = get-date
    $endTime   = $startTime.addMinutes($waitMinutes)
    $timeSpan = new-timespan $startTime $endTime
    
	write-host "`nWaiting OU Replication: $waitMinutes minutes..." -backgroundcolor black -foregroundcolor yellow
    while ($timeSpan -gt 0) {
        $timeSpan = new-timespan $(get-date) $endTime
        write-host "`r".padright(40," ") -nonewline
        write-host $([string]::Format("`rTime Remaining: {0:d2}:{1:d2}:{2:d2}", `
            $timeSpan.hours, `
            $timeSpan.minutes, `
            $timeSpan.seconds)) `
            -nonewline -backgroundcolor black -foregroundcolor yellow
        sleep 1
    }
	write-host ""
}

$FileName = Select-FileDialog -Title "Import an CSV file" -Directory "c:\"
$csvFile = Import-Csv $FileName
$OU = "OU=DistLists,OU=ExchangeUsers"

$domain = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
$DomainDN = (([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()).Domains | ? {$_.Name -eq $domain}).GetDirectoryEntry().distinguishedName
$final = "LDAP://$DomainDN"
$DomainPath = [ADSI]"$final"
$cOU = $DomainPath.Create("OrganizationalUnit",$OU)
$cOU.SetInfo()

CountDown 1.5

$OUPath = $OU+","+$DomainDN
$CreateOU = Get-OrganizationalUnit | where {$_.DistinguishedName -eq $OUPath}

foreach($entry in $csvFile){

	New-DistributionGroup -Alias $entry.samAccountName -DisplayName $entry.Name -Type Distribution -Name $entry.samAccountName -OrganizationalUnit $CreateOU

}
Write-Host "Script Completed" -ForegroundColor Yellow