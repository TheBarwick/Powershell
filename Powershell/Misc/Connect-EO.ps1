Function Connect-EO {
$host.ui.RawUI.WindowTitle = 'Exchange Online'
$Credentials = Get-Credential
$EOConnectionURI = "https://outlook.office365.com/powershell-liveid/"
$ExchangeOnlineSession = New-PSSession `
    -ConfigurationName Microsoft.Exchange `
    -ConnectionUri $EOConnectionURI `
    -Credential $Credentials `
    -Authentication "Basic" `
    -AllowRedirection
Import-PSSession $ExchangeOnlineSession 
}