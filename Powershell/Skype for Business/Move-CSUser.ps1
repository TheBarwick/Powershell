# Functions

# Strings
$console = $host.UI.RawUI
$console.BackgroundColor = "black"
$console.ForegroundColor = "gray"
$Credentials = Get-Credential

$URL = "https://admin1a.online.lync.com/HostedMigration/hostedmigrationservice.svc"
$Target = "sipfed.online.lync.com"
$CurrentDate = Get-Date -format "dd-MMM-yyyy HH:mm"
$CurrentDate = $CurrentDate.ToString().Replace(“:”, “-”)
$Logs = "e:\scripts\msonline\" + "Move-CSUser" + $CurrentDate + ".log"

Write-Host "Please enter the location of your list of users .csv file" -ForegroundColor "yellow"
$CSVLocation = Read-Host
$CSVData = Import-CSV $CSVLocation

ForEach ($Input in $CSVData) {
    $Identity = $Input.name
    Move-CsUser -Identity $Identity -Target $Target -Confirm:$False -Credential $Credentials -HostedMigrationOverrideUrl $URL -proxypool "pdcpool.mattel.com"
    $IdentityCheck = Get-CSUser $Input.name –Filter {HostingProvider –eq “sipfed.online.lync.com”}
    If ($IdentityCheck -) { 
        Write-Host "$Input.name has been successfully migrated from Skype for Business on-premise's to Skype for Business Online" -ForegroundColor "Green"
        Add-Content $Logs "$Input.name has been successfully migrated from Skype for Business on-premise's to Skype for Business Online"
    }
    Else {
        Write-Host "$Input.name has NOT been migrated from Skype for Business on-premise's to Skype for Business Online" –ForegroundColor “Dark Green"
        Add-Content $Logs "$Input.name has NOT been migrated from Skype for Business on-premise's to Skype for Business Online"
    }
}
