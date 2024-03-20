$OfficeUninstallStrings = ((Get-ItemProperty "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*") `
    + (Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*") | 
    Where {$_.DisplayName -like "*HP Connection Optimizer*" -or $_.DisplayName -like "*HP Documentation*"} | 
    Select UninstallString).UninstallString

    ForEach ($UninstallString in $OfficeUninstallStrings) {
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
    }

