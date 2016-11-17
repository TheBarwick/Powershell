#########################################################################################
# COMPANY: CDW                                                                          #
# NAME: FinishMoves.ps1                                                                 #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  03/25/2014                                                                     #
# EMAIL: Dean.SEsko@S3.CDW.com                                                          #
#                                                                                       #
# COMMENT:  Script to complete autosuspended mailbox moves in Office 365                #
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 08/28/2014 Initial Version.                                                       #
#                                                                                       #
#########################################################################################



cls

# Try To Connect to Online Tenant.  If fails make a connction

Try { $test = Get-MsolDomain -ErrorAction stop | where { ($_.name -like "*.mail.onmicrosoft.com") } }
Catch { Invoke-Expression .\Connect365.ps1 }
Finally { }


get-moverequest | Where {$_.status -eq "AutoSuspended"} | Resume-MoveRequest



