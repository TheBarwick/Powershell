#########################################################################################
# COMPANY: CDW                                                                          #
# NAME: Add-365UserLicense.ps1                                                          #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  03/25/2014                                                                     #
# EMAIL: Dean.SEsko@CDW.com                                                             #
#                                                                                       #
# COMMENT:  Script to Assign a License to an Office 365 User                            #
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 03/25/2014 Initial Version.                                                       #
# 1.1 10/17/2014 Foramt Cleanup and connection detection change.                        #
#                                                                                       #
#                                                                                       #
#########################################################################################
param ( [Parameter(Mandatory=$true)] 
  [string]$UPN
)
#	Setup Shell
$bg = (Get-Host).UI.RawUI
$BG.BackgroundColor = "black"
$RetryCount = 10
$index = 0

cls

# Try To Connect to Online Tenant.  If fails make a connction

Try {$test = Get-MsolDomain -ErrorAction stop | where {($_.name -like "*.mail.onmicrosoft.com")} }
Catch{Invoke-Expression .\Connect365.ps1 }
Finally{}

Try {$test=get-pssession -Name "ON-Prem-Exchange" -ErrorAction stop}
Catch{Invoke-Expression .\ConnectLocalExchange.ps1 }
Finally{}


Function Show-Menu {
cls
Write-host "Office 365 User Licenses"
Write-host ""
if ($OnlineUser.IsLicensed){  
    $LicString = $OnlineUser.Licenses.AccountSKUID.tostring()
    Write-host "$UPN is currently assigned $LicString" -ForegroundColor Green -BackgroundColor Black
    write-host
  }
Read-Host -Prompt $menudata
}



Function CheckOnlineUser{
    if($OnlineUser) {
        Set-MsolUser -UserPrincipalName $upn -UsageLocation "US"
        return $true
        }
    Else{
		Invoke-Expression .\StartSync.ps1
  	    Write-Host "Checking for Office 365 User:" $UPN -ForegroundColor Green -BackgroundColor Black
        sleep 30
        Return $false
    }
  }

if ($UPN){
 $user = get-onpremuser $upn -ErrorAction silentlycontinue
if ($user){   
    do {
        $OnlineUser=Get-MsolUser -UserPrincipalName $upn -ErrorAction silentlycontinue
        $GoodUser = CheckOnlineUser
        $index ++
        }
    until ($GoodUser -or ($index -ge $RetryCount ))

   
   if (!($index -ge $RetryCount )){
       
        $SKudata= Get-MsolAccountSku |Select-Object AccountSKUID,ActiveUnits,ConsumedUnits
        $skureport=@()
        $Report =@()
        $SkuIndex = 1
        if ($SKudata){
            foreach($sku in $SKUdata){
                $MyData = New-Object Object;
                $MyData | Add-Member NoteProperty "Index" $SkuIndex;
                $MyData | Add-Member NoteProperty "License SKU"  $sku.AccountSKUID
                $MyData | Add-Member NoteProperty "License Count"  $sku.ActiveUnits
                $MyData | Add-Member NoteProperty "Used Licenses"  $sku.ConsumedUnits
                $Report += $MyData 
                $SkuIndex++
            }
                $SkuReport += 
                $menudata= $Report |FT -Autosize |out-string
                $menudata += "Please Enter a License Number or Q for Quit"
                $exit = $false
                $365lic = ""
        
                Do{
            Switch (Show-Menu) {
                "1" {$365lic = $skudata[0].AccountSKUID
                    $exit = $true} 
                "2" {$365lic = $skudata[1].AccountSKUID
                    $exit = $true}
                "3" {$365lic = $skudata[2].AccountSKUID
                    $exit = $true} 
                "4" {$365lic = $skudata[3].AccountSKUID 
                    $exit = $true}  
                "5" {$365lic = $skudata[4].AccountSKUID
                    $exit = $true}
                "6" {$365lic = $skudata[5].AccountSKUID
                    $exit = $true} 
                "7" {$365lic = $skudata[6].AccountSKUID
                    $exit = $true}
                "8" {$365lic = $skudata[7].AccountSKUID
                    $exit = $true}    
                "9" {$365lic = $skudata[8].AccountSKUID
                    $exit = $true}
                "10"{$365lic = $skudata[9].AccountSKUID
                    $exit = $true}
                "11"{$365lic = $skudata[9].AccountSKUID
                    $exit = $true} 
                "12"{$365lic = $skudata[9].AccountSKUID
                    $exit = $true} 
                "13"{$365lic = $skudata[9].AccountSKUID
                    $exit = $true} 
                "14"{$365lic = $skudata[9].AccountSKUID
                    $exit = $true} 
                "15"{$365lic = $skudata[9].AccountSKUID
                    $exit = $true} 
                "Q" {$365lic = "NA"
                    $exit = $true}
                default { $exit = $false
                    $365lic = ""  }
                    }

           }

                While ( ($365lic.length -eq 0) -and ($exit -ne $true ) -or  ($365lic.length -eq ""))

                if (($365lic -ne "NA")-and ($365lic.count -ne "0")){
                    if (!($OnlineUser.IsLicensed)){  
                    Try {
                         Set-MsolUserLicense -UserPrincipalName $UPN  -AddLicenses $365lic 
                         Write-host ""
                         Write-host ""
                         write-Host "License Assigned" -ForegroundColor Green -BackgroundColor black
                         Write-host ""
                         Write-host ""    }
                    Catch{
                          Write-host ""
                          Write-host ""
                          Write-Host "Error assigning license to user" -ForegroundColor Red -BackgroundColor Black
                          Write-host ""
                          Write-host ""
                          }
                    Finally{}
                    }
                    Else{

                         if ($OnlineUser.Licenses.AccountSKUID -eq $365lic){
                            Try {
                                $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $365lic 
                                Set-MsolUserLicense -UserPrincipalName $UPN -LicenseOptions $LicenseOptions 
                                Write-host ""
                                Write-host ""
                                write-Host "License Re-Assigned" -ForegroundColor Green -BackgroundColor black
                                Write-host ""
                                Write-host ""    }
                            Catch{
                                Write-host ""
                                Write-host ""
                                Write-Host "Error assigning license to user" -ForegroundColor Red -BackgroundColor Black
                                Write-host ""
                                Write-host ""
                                }
                            Finally{}
                         }
                         Else{
                          Try {
                               
                                Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses  $OnlineUser.Licenses.AccountSKUID  -AddLicenses $365lic 
                                $LicString = $OnlineUser.Licenses.AccountSKUID.tostring()
                                Write-host ""
                                Write-host ""
                                write-Host "License Switched from $LicString to $365lic " -ForegroundColor Green -BackgroundColor black
                                Write-host ""
                                Write-host ""    }
                            Catch{
                                Write-host ""
                                Write-host ""
                                Write-Host "Error assigning license to user" -ForegroundColor Red -BackgroundColor Black
                                Write-host ""
                                Write-host ""
                                }
                            Finally{}
                    }
                    }
                   
                   }
                
                Else{
                        cls
                        Write-host
                        Write-host "No License Selected" -ForegroundColor Yellow -BackgroundColor Black
                        Write-host
                        Write-host
                      }

     }
     
}
}
Else{
    CLS
    Write-Host""
    Write-Host""
	Write-Host "User Does Not Exist." -ForegroundColor Red -BackgroundColor Black
    Write-Host""
    Write-Host""	 
}
 
}