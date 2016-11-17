$Credentials = Get-Credential
Import-Module MSOnline
Connect-MSOLService -Credential $Credentials
$MSOLAccountSKU = Get-MsolAccountSku
$TenantID = $MSOLAccountSku.AccountObjectId
$TenantID