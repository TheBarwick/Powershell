$location    = "West US"
$serviceName = "mpn100west"
$vmName      = "UbuntuVM01"
$size        = "Small"
$adminUser   = "tcvadmin"
$password    = "Secure01"

$imageName = 'b39f27a8b8c64d52b05eac6a62ebad85__Ubuntu_DAILY_BUILD-wily-15_10-amd64-server-20160622-en-us-30GB'

$certPath    = "D:\OneDrive\Home Lab Resources\Azure Resources - thecloudvisor.com\mpn100west_Cert.pem"

New-AzureService -ServiceName $serviceName `
                 -Location $location

$cert = Get-PfxCertificate -FilePath $certPath

Add-AzureCertificate -CertToDeploy $certPath `
                     -ServiceName $serviceName

$sshKey = New-AzureSSHKey -PublicKey -Fingerprint $cert.Thumbprint `
                          -Path "/home/$linuxUser/.ssh/authorized_keys"

New-AzureVMConfig -Name $vmName -InstanceSize $size -ImageName $imageName ` |
Add-AzureProvisioningConfig -Linux -LinuxUser $adminUser -Password $password -SSHPublicKeys $sshKey ` |
New-AzureVM -ServiceName $serviceName