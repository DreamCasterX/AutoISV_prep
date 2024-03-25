$OfficeUninstallStrings = ((Get-ItemProperty "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*") `
    + (Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*") + (Get-ItemProperty "Registry::HKEY_USERS\S-1-5-21-3638432799-4170261345-3180263834-1001\Software\Microsoft\Windows\CurrentVersion\Uninstall\*") | 
    Where {$_.DisplayName -like "*Microsoft 365 -*" -or $_.DisplayName -like "*Microsoft OneNote -*" -or $_.DisplayName -like "*OneDrive*"} | 
    Select UninstallString).UninstallString

    ForEach ($UninstallString in $OfficeUninstallStrings) {
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
    }

