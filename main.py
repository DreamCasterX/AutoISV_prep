import subprocess
import sys
import shutil
import os


def ask(func):
    def inner(*args, **kwargs):
        func(*args, **kwargs)
        while True:
            answer = input("\nMove to the next case? (y/n)\n")
            if answer.lower() == "y":
                break
            if answer.lower() == "n":
                sys.exit()
            else:
                continue

    return inner


print("\nStarting ISV system preperation auto tool...\n\n")


# Check if any devices reports error status (# 1)
@ask
def case_01():
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
            "#1 - YB found in device manager!! Please look into the problematic device(s):\n"
        )
        for line in output_lines:
            print(line)
        while True:
            confirm = input("\nDo you want to continue? (y/n) ")
            if confirm.lower() == "y":
                break
            if confirm.lower() == "n":
                sys.exit()
            else:
                continue
    else:
        print("#1 - Ensure no YB found in device manager [Complete]")


# Check Intel DPTF support (#2)
@ask
def case_02():
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
@ask
def case_03():
    get_chassis = subprocess.run(
        ["wmic", "systemenclosure", "get", "chassistypes"],
        stdout=subprocess.PIPE,
        text=True,
    )
    chassis_string = get_chassis.stdout.strip()
    power_plan_all = subprocess.run(
        [
            "powercfg",
            "/list",
        ],
        stdout=subprocess.PIPE,
        text=True,
    )
    power_plan_all_string = power_plan_all.strip()
    power_plan_current = subprocess.run(
        [
            "powercfg",
            "/GetActiveScheme",
        ],
        stdout=subprocess.PIPE,
        text=True,
    )
    power_plan_current_string = power_plan_current.strip()

    if "{10}" in chassis_string:  # Notebook
        if "HP Optimized" in power_plan_all_string:
            # # Get GUID of HP Optimized
            # for i in
            # guid =

            # scheme_guid = str(subprocess.check_output(["powercfg", "-getactivescheme"]))
            # current_scheme_guid = scheme_guid[scheme_guid.index("GUID: "):][6:42]
            # # current_scheme_guid = scheme_guid[-49:-13]
            # # print(scheme_guid)
            # # print(current_scheme_guid)
            # sub_guid = str(subprocess.check_output(["powercfg", "-aliases"]))
            # sleep_guid = sub_guid[:sub_guid.index("  SUB_SLEEP")][-36:]
            # # print(sleep_guid)
            # output = str(subprocess.check_output(["powercfg", "-query", current_scheme_guid, sleep_guid]))
            # # print(output)
            # ac_output = output[output.index("STANDBYIDLE"):][202:244]
            # # print(ac_output)

            if "HP Optimized" in power_plan_current_string:
                print("#3 - <NB> Power plan is set to HP Optimized [Complete]")
            if "HP Optimized" not in power_plan_current_string:
                print(
                    "#3 - <NB> Power plan is not set to HP Optimized !!"
                )  # Uncomment if fix this
                # TODO: Set power plan to HP Optimized
                # Get GUID of HP Optimized power plan and run powercfg /setactive {GUID}
                print("#3 - <NB> Power plan is set to HP Optimized [Complete]")
        if "HP Optimized" not in power_plan_all_string:
            if "Balanced" in power_plan_current_string:
                print(
                    "#3 - <NB> Power plan for HP Optimized is not availalbe. Use Balanced [Complete]"
                )
            else:
                print(
                    "#3 - <NB> Power plan for HP Optimized is not availalbe. Current power plan is not set to Balanced!!"
                )
                sys.exit()
    else:  # Desktop
        if "High performance" in power_plan_all_string:
            if "High performance" in power_plan_current_string:
                print(
                    "#3 - <DT> Current power plan is set to High performance [Complete]"
                )
            if "High performance" not in power_plan_current_string:
                print(
                    "#3 - <DT> Current power plan is not set to High performance !!"
                )  # Uncomment if fix this
                # TODO: Set power plan to High performance
                # Get GUID of High performance power plan and run powercfg /setactive {GUID}
                print(
                    "#3 - <DT> Current power plan is set to High performance [Complete]"
                )
        if "High performance" not in power_plan_all_string:
            print("#3 - <DT> Power plan for High performance is not availalbe!!")
            sys.exit()


# TODO: Turn off Wi-Fi and turn on airplane mode (#4)  *Reboot required
@ask
def case_04():
    subprocess.run(  # Turn off Wi-Fi
        [
            "powershell",
            'Set-NetAdapterAdvancedProperty -Name "Wi-Fi" -AllProperties ',
            '-RegistryKeyword "SoftwareRadioOff" -RegistryValue "1"',
        ],
        check=True,
    )
    subprocess.run(  # TODO: Turn on airplane mode (fail)
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
@ask
def case_05():
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
@ask
def case_06():
    reg_file_path = "./src/Rt Click Options.reg"
    try:
        os.system(f'reg import \"{reg_file_path}\" > nul 2>&1')  # fmt: skip
        print("#6 - Add RT Click Options regestry [Complete]")
    except:
        print("#6 - Failed to add RT Click Options regestry!!")
        sys.exit()


# Check if Secure Boot is disabled and enable test signing (# 7)
@ask
def case_07():
    result = subprocess.run(
        ["powershell", "Confirm-SecureBootUEFI"], stdout=subprocess.PIPE, text=True
    )
    sb_state = result.stdout.strip()
    if sb_state.lower() == "true":
        print("Secure boot is enabled!! Enter BIOS to disable Secure boot first")
        sys.exit()
    else:
        try:
            subprocess.run(
                ["bcdedit", "/set", "testsigning", "on"],
                check=True,
                stdout=subprocess.DEVNULL,
            )
            print("#7 - Enable test mode [Complete]")
        except:
            print("#7 - Failed to enable test mode!!")
            sys.exit()


# Copy Power Config folder and import power scheme (#8)
@ask
def case_08():
    source_folder = "./src/PowerConfig"
    destination_folder = "C:\\PowerConfig"
    try:
        if os.path.exists(destination_folder):
            shutil.rmtree(destination_folder)
        shutil.copytree(source_folder, destination_folder)
        subprocess.run(
            ["C:\\PowerConfig\\Install.bat"],
            check=True,
            stdout=subprocess.DEVNULL,  # Uncomment to see power scheme GUID
        )
        print("#8 - Copy PowerConfig folder and import power scheme [Complete]")
    except:
        print("#8 - Failed to copy PowerConfig folder and import power scheme!!")
        sys.exit()


# Display full path/hidden files/empty drives/extensions/merge conflicts/protected OS files (#9)
@ask
def case_09():
    subprocess.run(  # Display the full path in the title bar
        [
            "reg",
            "add",
            r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CabinetState",
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
    subprocess.run(  # Show hidden files. folders. and drives
        [
            "reg",
            "add",
            r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced",
            "/v",
            "Hidden",
            "/t",
            "REG_DWORD",
            "/d",
            "1",
            "/f",
        ],
        check=True,
        stdout=subprocess.DEVNULL,
    )
    subprocess.run(  # Show empty drives
        [
            "reg",
            "add",
            r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced",
            "/v",
            "HideDrivesWithNoMedia",
            "/t",
            "REG_DWORD",
            "/d",
            "0",
            "/f",
        ],
        check=True,
        stdout=subprocess.DEVNULL,
    )
    subprocess.run(  # Show extensions for knowm file types
        [
            "reg",
            "add",
            r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced",
            "/v",
            "HideFileExt",
            "/t",
            "REG_DWORD",
            "/d",
            "0",
            "/f",
        ],
        check=True,
        stdout=subprocess.DEVNULL,
    )
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
    subprocess.run(  # Show protected OS files
        [
            "reg",
            "add",
            r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced",
            "/v",
            "ShowSuperHidden",
            "/t",
            "REG_DWORD",
            "/d",
            "1",
            "/f",
        ],
        check=True,
        stdout=subprocess.DEVNULL,
    )
    subprocess.run(  # Restart Windows Explorer to take effect changes immediately
        [
            "taskkill",
            "/F",
            "/IM",
            "explorer.exe",
            "&",
            "start",
            "explorer",
        ],
        check=True,
        stdout=subprocess.DEVNULL,
    )
    print(
        "#9 - Display full path/hidden files/empty drives/extensions/merge conflicts/protected OS files [Complete]"
    )


# Set sleep & display off to Never in power option (#10)
@ask
def case_10():
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
@ask
def case_11():
    subprocess.run(
        ["tzutil", "/s", "Central Standard Time"],
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
    print(
        "#11 - Set time zone to Central US and disable set time automatically [Complete]"
    )


# Auto hide the taskbar (#12)
@ask
def case_12():
    subprocess.run(
        [
            "powershell",
            "-Command",
            "&{$p='HKCU:SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\StuckRects3';$v=(Get-ItemProperty -Path $p).Settings;$v[8]=3;&Set-ItemProperty -Path $p -Name Settings -Value $v;&Stop-Process -f -ProcessName explorer}",
        ],
        check=True,
    )
    print("#12 - Auto hide the taskbar [Complete]")


# Unpin Edge and pin Paint/Snipping Tool to taskbar (#13)
@ask
def case_13():
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
        [
            reg_file_path,
            f"{app_path}\\Microsoft\\WindowsApps\\SnippingTool.exe",
            "5386",
        ],
        check=True,
        stdout=subprocess.DEVNULL,
    )
    os.system("taskkill /F /IM explorer.exe > nul 2>&1")  # Restart Windows Explorer
    os.system("start explorer.exe > nul 2>&1")
    print("#13 - Unpin Edge and pin Paint/Snipping Tool to taskbar [Complete]")


# TODO: Disable UAC prompt (# 14)  需要重開機
# @ask
# def case_14():


# Turn off Windows Defender Firewall (# 15)
@ask
def case_15():
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
@ask
def case_16():
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
@ask
def case_18():
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
    print(
        "#18 - Set Brightness level to 100% and disable adaptive brightness [Complete]"
    )


# Uninstall MS Office (#19)
# TODO: Remove MS Office 365/One Note/Teams from Installed apps
@ask
def case_19():
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
@ask
def case_25():
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
def case_28():
    subprocess.run(
        ["powershell", "-Command", "powercfg -h off"],
        check=True,
    )
    print("#28 - Disable system hibernation [Complete]")


# Run case
case_01()
case_02()
case_03()
case_04()
case_05()
case_06()
case_07()
case_08()
case_09()
case_10()
case_11()
case_12()
case_13()
# case_14()
case_15()
case_16()
# case_17()
case_18()
case_19()
# case_20()
# case_21()
# case_22()
# case_23()
# case_24()
case_25()
# case_26()
# case_27()
case_28()
