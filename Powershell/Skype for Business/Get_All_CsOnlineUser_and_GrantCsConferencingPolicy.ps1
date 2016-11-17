$Users = Get-CsOnlineUser
ForEach ($User in $Users) {
$User = $User.UserPrincipalName
$GrantPolicy = Grant-CsConferencingPolicy -Identity $User -PolicyName ""
$GrantPolicy
}