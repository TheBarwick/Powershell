#
# Set_Inheritance.ps1
#

Import-Module ActiveDirectory

Function Set-Inheritance { 
    param($ObjectPath) 
    $ACL = Get-ACL -path “AD:\$ObjectPath” 
    If ($acl.AreAccessRulesProtected){ 
         $ACL.SetAccessRuleProtection($False, $True) 
        Set-ACL -AclObject $ACL -path “AD:\$ObjectPath” 
        Write-Host “Updated: “$ObjectPath 
    } # Close IF 
} # Close Set-Inheritance

# Find user with AdminCount set to 1 (Inheritance Disabled)
$users = get-aduser -SearchBase “OU=Users,OU=Accounts,OU=TheBarwick,DC=Thebarwick,DC=com” -Filter {AdminCount -eq 1} 

# Enable inheritance flag for each user 
$users | foreach {Set-Inheritance $_.distinguishedname}