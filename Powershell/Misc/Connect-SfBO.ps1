Function Connect-SfBO {
    $Credentials = Get-Credential
    $host.ui.RawUI.WindowTitle = 'Skype for Business Online'
    Import-Module SkypeOnlineConnector
    $SfBOSession = New-CsOnlineSession `
        -Credential $Credentials
    Import-PSSession $SfBOSession
}