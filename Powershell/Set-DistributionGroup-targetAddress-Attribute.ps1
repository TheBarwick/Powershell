$CSV = import-csv "C:\CDW_GAC_Migration\CSV\dl.csv"
$CSV | foreach {
$SMTP = "SMTP:" 
$Email = $SMTP + $_.EMAILADDRESS
set-ADGroup -identity $_.DISPLAYNAME -add @{targetAddress= $Email}
}