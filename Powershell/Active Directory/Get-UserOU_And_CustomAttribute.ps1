Get-Mailbox * | Select-Object SamAccountName,CustomAttribute5 | Export-Csv "C:\Get-Mailboxes.csv" -NoTypeInformation

$CSVData = Import-CSV "C:\Get-Mailboxes.csv"
$CSVData | ForEach {Get-ADUser -Identity $_.SamAccountName -Properties * | Select-Object @{n='OU';e={$_.canonicalname -replace "/$($_.cn)",""}}} | Export-Csv "C:\Get-ADUsers.csv" -Encoding "Unicode" -NoTypeInformation
$MailboxesFile = 'C:\Get-Mailboxes.csv'
$ADUsersFile = 'C:\Get-ADUsers.csv'
$DestinationFile = 'C:\UserInfo.csv'
@(Import-Csv $MailboxesFile) + @(Import-Csv $ADUsersFile) | Export-Csv $DestinationFile -NoTypeInformation