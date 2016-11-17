Import-Module ActiveDirectory
Get-ADGroup -Filter * -SearchBase "OU=Exchange Distribution Groups,OU=Groups,OU=GAC,OU=Migration,D
C=rayovac,DC=com" | Set-ADGroup -Clear targetAddress