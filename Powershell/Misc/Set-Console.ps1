    Function Set-Console {
        # Forces PowerShell Console to run with Administrative Rights
        If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")){   
            $Arguments = "& '" + $MyInvocation.mycommand.definition + "'"
            Start-Process powershell `
                -Verb runAs `
                -ArgumentList $arguments
            Break
         }
    }