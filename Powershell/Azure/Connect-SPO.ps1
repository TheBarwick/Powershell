Function Connect-SPO {
    $Credentials = Get-Credential
    $host.ui.RawUI.WindowTitle = 'Sharepoint Online'
    $SPOModule = Microsoft.Online.SharePoint.PowerShell
    $SPOConnectionURL = "https://domainhost-admin.sharepoint.com"
        Import-Module $SPOModule ` -DisableNameChecking
        Connect-SPOService `
            -Url $SPOConnectionURL `
            -Credential $Credentials
}