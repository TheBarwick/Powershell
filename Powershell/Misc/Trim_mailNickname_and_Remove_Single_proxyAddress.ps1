Import-Module ActiveDirectory
Get-ADObject -SearchBase "OU=contactsDC=domain,DC=com" -LDAPFilter "objectClass=contact" -Properties Name, ObjectGUID, mailNickname | %{Set-ADObject -Identity $_.ObjectGUID -Replace @{mailNickname=$_.Name.split('@')[0]}}

$OU = [ADSI]"LDAP://OU=Groups,OU=Z_Notes_Contacts,DC=domain,DC=com"
ForEach ($Contact in $OU.PsBase.Children) {
  $Contact.Get("proxyAddresses") | ?{ $_ -Like "*@notes.domain.com" } | %{
    $Contact.PutEx(4, "proxyAddresses", @("$_"))
    $Contact.SetInfo()
  }
}

$OU = [ADSI]"LDAP://OU=Resources,OU=Z_Notes_Contacts,DC=domain,DC=com"
ForEach ($Contact in $OU.PsBase.Children) {
  $Contact.Get("proxyAddresses") | ?{ $_ -Like "*@notes.domain.com" } | %{
    $Contact.PutEx(4, "proxyAddresses", @("$_"))
    $Contact.SetInfo()
  }
}

$OU = [ADSI]"LDAP://OU=Users_Contacts,OU=Z_Notes_Contacts,DC=domain,DC=com"
ForEach ($Contact in $OU.PsBase.Children) {
  $Contact.Get("proxyAddresses") | ?{ $_ -Like "*@notes.domain.com" } | %{
    $Contact.PutEx(4, "proxyAddresses", @("$_"))
    $Contact.SetInfo()
  }
}