 Function Connect-AzureRM {
    $Credentials = Get-Credential
    Import-Module Azure
    Login-AzureRmAccount `
        -Credential $Credentials
}
