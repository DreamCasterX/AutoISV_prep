import subprocess
import sys
import shutil
import os


# Check if any devices reports error status (# 1)
result = subprocess.run(
    [
        "powershell",
        "-command",
        "(Get-PnpDevice) | Where-Object { $_.Status -eq 'Error' }",
    ],
    stdout=subprocess.PIPE,
    text=True,
)
output_lines = result.stdout.strip().splitlines()
if output_lines:
    print(
        "#1 - YB found in device manager!! Please look into the problematic device(s):"
    )
    for line in output_lines:
        print(line)
    while True:
        answer = input("\nDo you want to continue? (y/n)")
        if answer.lower() == "y":
            break
        if answer.lower() == "n":
            sys.exit()
        else:
            continue
else:
    print("#1 - Ensure no YB found in device manager [Complete]")


# Check Intel DPTF support (#2)
result = subprocess.run(
    ["wmic", "cpu", "get", "caption"],
    stdout=subprocess.PIPE,
    text=True,
)
cpu_arch = result.stdout.strip()
if "AMD64" in cpu_arch:
    print("#2 - Skip checking Intel DPTF on AMD platform [Complete]")
if "Intel64" in cpu_arch:
    result = subprocess.run(
        [
            "powershell",
            "-command",
            "(Get-PnpDevice) | Where-Object { $_.FriendlyName -like '*Intel(R) Dynamic Tuning*' }",
        ],
        stdout=subprocess.PIPE,
        text=True,
    )
    output_string = result.stdout.strip()
    if "Unknown" in output_string:
        print(
            "#2 - Intel DPTF is supported but NOT enabled!! Please enter BIOS to enable it"
        )
        sys.exit()
    if "OK" in output_string:
        print("#2 - Intel DPTF is already enabled [Complete]")
    else:
        print("#2 - Intel DPTF is not supported on this platform [Complete]")


# TODO: Get the chassis type and set corresponding power plan (#3)
result = subprocess.run(
    ["wmic", "systemenclosure", "get", "chassistypes"],
    stdout=subprocess.PIPE,
    text=True,
)
chassis_type = result.stdout.strip()
if "{10}" in chassis_type:  # Notebook
    print("#3 - This is a Notebook")
else:  # Desktop/AIO
    print("#3 - This is NOT a Notebook")


# TODO: Turn off Wi-Fi and turn on airplane mode (#4)  *Reboot required
subprocess.run(  # Turn off Wi-Fi
    [
        "powershell",
        'Set-NetAdapterAdvancedProperty -Name "Wi-Fi" -AllProperties ',
        '-RegistryKeyword "SoftwareRadioOff" -RegistryValue "1"',
    ],
    check=True,
)
subprocess.run(  # Turn on airplane mode (fail)
    [
        "reg",
        "add",
        r"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\RadioManagement",
        "/v",
        "SystemRadioState",
        "/t",
        "REG_DWORD",
        "/d",
        "1",
        "/f",
    ],
    check=True,
    stdout=subprocess.DEVNULL,
)
print("#4 - Turn off Wi-Fi and turn on airplane mode [Complete]")


# Turn off User Account Control (UAC) (#5)
subprocess.run(
    [
        "reg",
        "add",
        r"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\policies\system",
        "/v",
        "EnableLUA",
        "/t",
        "REG_DWORD",
        "/d",
        "0",
        "/f",
    ],
    check=True,
    stdout=subprocess.DEVNULL,
)
print("#5 - Turn off User Account Control (UAC) [Complete]")


# Add RT Click Options regestry (#6)
reg_file_path = "./src/Rt Click Options.reg"
try:
    subprocess.run(
        ["reg.exe", "import", reg_file_path],
        check=True,
        stdout=subprocess.DEVNULL,
    )
    print("#6 - Add 'RT Click Options' regestry [Complete]")
except subprocess.CalledProcessError:
    print("#6 - Failed to add 'RT Click Options' regestry!! Make sure the file exists")
    sys.exit()


# Check if Secure Boot is disabled and enable test signing (# 7)
result = subprocess.run(
    ["powershell", "Confirm-SecureBootUEFI"], stdout=subprocess.PIPE, text=True
)
sb_state = result.stdout.strip()
if sb_state.lower() == "true":
    print("Secure boot is enabled!! Enter BIOS to disable Secure boot first")
    sys.exit()
else:
    subprocess.run(
        ["bcdedit", "/set", "testsigning", "on"],
        check=True,
        stdout=subprocess.DEVNULL,
    )
    print("#7 - Enable test mode [Complete]")


# Copy Power Config folder and import power scheme (#8)
source_folder = "./src/PowerConfig"
destination_folder = "C:\\PowerConfig"
if os.path.exists(destination_folder):
    shutil.rmtree(destination_folder)
shutil.copytree(source_folder, destination_folder)
subprocess.run(
    ["C:\\PowerConfig\\Install.bat"],
    check=True,
    stdout=subprocess.DEVNULL,
)
print("#8 - Copy PowerConfig folder and import power scheme [Complete]")


# TODO: Display the full path/Show Hidden Files and uncheck Hide empty drives, uncheck Hide extensions,
# uncheck Hide folders merge conflicts, uncheck Hide Protected OS files  (#9)
subprocess.run(  # Show full path
    [
        "reg",
        "add",
        r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\\Explorer\CabinetState",
        "/v",
        "FullPath",
        "/t",
        "REG_DWORD",
        "/d",
        "1",
        "/f",
    ],
    check=True,
    stdout=subprocess.DEVNULL,
)
# Show Hidden Files (FAIL)
# reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "HideFileExt" /t REG_DWORD /d 0 /f

# Show empty drives (FAIL)
# reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "NoDrives" /t REG_DWORD /d 0 /f

subprocess.run(  # Show folders merge conflicts
    [
        "reg",
        "add",
        r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced",
        "/v",
        "HideMergeConflicts",
        "/t",
        "REG_DWORD",
        "/d",
        "0",
        "/f",
    ],
    check=True,
    stdout=subprocess.DEVNULL,
)
# Show protected OS files (FAIL)
# reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "Hidden" /t REG_DWORD /d 0 /f
print(
    "#9 - Display full path and show hidden folder/files/empty drives/extensions [Complete]"
)


# Set sleep & display off to Never in power option (#10)
subprocess.run(
    [
        "powercfg",
        "/change",
        "standby-timeout-ac",
        "0",
    ],
    check=True,
)
subprocess.run(
    [
        "powercfg",
        "/change",
        "standby-timeout-dc",
        "0",
    ],
    check=True,
)
subprocess.run(
    [
        "powercfg",
        "/change",
        "monitor-timeout-ac",
        "0",
    ],
    check=True,
)
subprocess.run(
    [
        "powercfg",
        "/change",
        "monitor-timeout-dc",
        "0",
    ],
    check=True,
)
print("#10 - Set sleep & display off to Never in power option [Complete]")


# Set time zone to Central US and disable auto set time (# 11)
subprocess.run(
    ["tzutil", "/s", '"Central Standard Time"'],
    check=True,
)
subprocess.run(
    [
        "powershell",
        "-Command",
        'Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Services\\W32Time\\Parameters" -Name "Type" -Value "NoSync"',
    ],
    check=True,
)
print("#11 - Set time zone to Central US and disable set time automatically [Complete]")


# Auto hide the taskbar (#12)
subprocess.run(
    [
        "powershell",
        "-Command",
        "&{$p='HKCU:SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\StuckRects3';$v=(Get-ItemProperty -Path $p).Settings;$v[8]=3;&Set-ItemProperty -Path $p -Name Settings -Value $v;&Stop-Process -f -ProcessName explorer}",
    ],
    check=True,
)


# Unpin Edge and pin Paint/Snipping Tool to taskbar (#13)
reg_file_path = "./src/syspin.exe"
app_path = os.environ.get("LocalAppData")
subprocess.run(  # Unping Edge
    [
        "powershell",
        "-command",
        "Remove-Item 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Taskband' -Recurse -Force",
    ],
    check=True,
)
subprocess.run(  # Ping Paint
    [reg_file_path, f"{app_path}\\Microsoft\\WindowsApps\\mspaint.exe", "5386"],
    check=True,
    stdout=subprocess.DEVNULL,
)
subprocess.run(  # Ping Snipping Tool
    [reg_file_path, f"{app_path}\\Microsoft\\WindowsApps\\SnippingTool.exe", "5386"],
    check=True,
    stdout=subprocess.DEVNULL,
)
# subprocess.run(  # Restart the Explorer
#     [
#         "powershell",
#         "-command",
#         "Stop-Process -Processname Explorer -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Force",
#     ],
#     check=True,
# )
print("#13 - Unpin Edge and pin Paint/Snipping Tool to taskbar [Complete]")


# TODO: Disable UAC prompt (# 14)  需要重開機


# Turn off Windows Defender Firewall (# 15)
subprocess.run(
    ["netsh", "advfirewall", "set", "domianprofile", "state", "off"],
    check=True,
)
subprocess.run(
    ["netsh", "advfirewall", "set", "privateprofile", "state", "off"],
    check=True,
)
subprocess.run(
    ["netsh", "advfirewall", "set", "publicprofile", "state", "off"],
    check=True,
)
print("#15 - Disable Windows Firewall Defender [Complete]")


# Set resolution to 1920x1080 and DPI to 100% (#17)
res_app_path = "./src/QRes.exe"
dpi_app_path = "./src/SetDpi.exe"
subprocess.run(
    [res_app_path, "/x:1920", "/y:1080"],
    check=True,
    stdout=subprocess.DEVNULL,
)
subprocess.run(
    [dpi_app_path, "100"],
    check=True,
    stdout=subprocess.DEVNULL,
)
print("#17 - Set resolution to 1920x1080 @ 100% [Complete]")


# Set brightness level to 100% and disable adaptive brightness (# 18)
subprocess.run(
    [
        "powershell",
        "-command",
        "(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,100)",
    ],
    check=True,
)
subprocess.run(
    [
        "powercfg",
        "-setacvalueindex",
        "SCHEME_CURRENT",
        "7516b95f-f776-4464-8c53-06167f40cc99",
        "FBD9AA66-9553-4097-BA44-ED6E9D65EAB8",
        "0",
    ],
    check=True,
)
subprocess.run(
    [
        "powercfg",
        "-setdcvalueindex",
        "SCHEME_CURRENT",
        "7516b95f-f776-4464-8c53-06167f40cc99",
        "FBD9AA66-9553-4097-BA44-ED6E9D65EAB8",
        "0",
    ],
    check=True,
)
subprocess.run(["powercfg", "-SetActive", "SCHEME_CURRENT"], check=True)
print("#18 - Set Brightness level to 100% and disable adaptive brightness [Complete]")

# Uninstall MS Office (#19)
# TODO: Remove MS Office 365/One Note/Teams from Installed apps
subprocess.run(
    [
        "powershell",
        "-Command",
        "Get-AppxPackage *OfficeHub* | Remove-AppxPackage; Get-AppxPackage *OneNote* | Remove-AppxPackage; Get-AppxPackage *Office* | Remove-AppxPackage",
    ],
    shell=True,
    check=True,
)
print("#19 - Uninsalled MS Office [Complete]")


# TODO: Uninstall HP apps (#20)

# TODO: Install .NET Framwork 3.5 (#21)

# TODO: Pause Windows Update and disable Allow downloads from other PCs (#22)


# Set regristry for DriverSearching to 0  (#25)
subprocess.run(
    [
        "reg",
        "add",
        r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\DriverSearching",
        "/v",
        "SearchOrderConfig",
        "/t",
        "REG_DWORD",
        "/d",
        "0",
        "/f",
    ],
    check=True,
    stdout=subprocess.DEVNULL,
)
print("#25 - Set Registry key for DriverSearching to 0 [Complete]")

# TODO: Turn off Smart app control/Reputation-based protection/Isolated browsing (#26)

# Stop WU service (#27)

# net stop wuauserv > NUL
# net stop bits > NUL
# net stop dosvc > NUL
# echo (4) WU is stopped
# echo.


# Disable system hibernation (#28)
subprocess.run(
    ["powershell", "-Command", "powercfg -h off"],
    check=True,
)
print("#28 - Disable system hibernation [Complete]")
