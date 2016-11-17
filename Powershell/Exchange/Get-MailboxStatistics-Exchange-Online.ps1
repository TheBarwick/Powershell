$securepassword = ConvertTo-SecureString -string '$CDW20151001$' -AsPlainText -Force 
$EOCredential = new-object System.Management.Automation.PSCredential ("brad.stevens@armoredautogroup.com", $securepassword)
$EOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $EOCredential -Authentication Basic -AllowRedirection
Import-PSSession $EOSession

$CSVData = Import-CSV "C:\CDW_GAC_Migration\CSV\Users.csv"
$CSVData | ForEach {Get-MailboxStatistics -Identity  $_.SamAccountName | Select-Object DisplayName,TotalItemSize,ItemCount,TotalDeletedItemSize,DeletedItemCount } | Export-CSV "C:\CDW_GAC_Migration\CSV\AAG-MailboxStatisticsExport.csv" -NoTypeInformation
Get-PSSession | Remove-PSSession
