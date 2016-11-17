Function Connect-MSOL {
    $host.ui.RawUI.WindowTitle = 'MSOL Online'
    Import-Module MsOnline
    Connect-MsolService `
        -Credential $Credentials
}