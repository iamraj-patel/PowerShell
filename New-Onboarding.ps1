$Host.UI.RawUI.FlushInputBuffer()

# Setting location for script so it can go to Automate folder and install application
try {
    Set-Location -Path $PSScriptRoot
    $make = (Get-CimInstance Win32_ComputerSystem).Manufacturer
}
catch {
    Write-Host "Error setting location or getting computer make: $($_.Exception.Message)"
}

# Create a log file
#$logFile = "$env:TEMP\script_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$logFile = ".\Log\Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
Start-Transcript -Path $logFile -Append

# Syncing time
try {
    net stop w32time
    w32tm /unregister
    w32tm /register
    net start w32time
    w32tm /resync
}
catch {
    Write-Host "Error syncing time: $($_.Exception.Message)"
}

# Installing Automate
try {
    $currentdir = $pwd
    $hostname = hostname
    $path = ".\Automate"
    cd $path
    Start-Process msiexec.exe -Wait -ArgumentList '/i "Agent_Install.msi" /quiet'
    Write-Host "Installation of Automate on $hostname has been finished."
    cd $currentdir
}
catch {
    Write-Host "Error installing Automate: $($_.Exception.Message)"
}

# Checking updates in Dell command update or Lenovo vantage or Lenovo Commercial Vantage.
try {
    if ($make -eq "Dell Inc.") {
        # Checking updates for Dell Drivers and Installing it.
        $DownloadLocation = "C:\\Program Files (x86)\\Dell\\CommandUpdate"
        Start-Process "$($DownloadLocation)\\dcu-cli.exe" -ArgumentList "/applyUpdates -autoSuspendBitLocker=enable -reboot=enable" -Wait
    }
    elseif ($make -eq "LENOVO") {
        Write-Host "We don't have any Automated ways to install Firmwares and Drivers for your system. Please update Drivers, Firmwares, and BIOS using Lenovo Commercial Vantage or Lenovo Vantage."
    }
    else {
        Write-Host "No make was selected."
    }
}
catch {
    Write-Host "Error checking updates: $($_.Exception.Message)"
}

# Install Windows Updates
try {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$False
    Install-Module PSWindowsUpdate -Confirm:$False -Force
    Import-Module PSWindowsUpdate
    Get-WindowsUpdate -Install -AcceptAll -AutoReboot
}
catch {
    Write-Host "Error installing Windows updates: $($_.Exception.Message)"
}

# Setting up password for ITADMIN
try {
    net user itadmin itadmin
}
catch {
    Write-Host "Error setting up ITADMIN account: $($_.Exception.Message)"
}

# Disable Administrator account
try {
    $adminAccount = Get-WmiObject Win32_UserAccount -Filter "Name='Administrator'"
    $adminAccount.Disabled = $true
    $adminAccount.Put()
}
catch {
    Write-Host "Error disabling Administrator account: $($_.Exception.Message)"
}

# Changing execution Policy from Bypass to Restricted.
try {
    Set-ExecutionPolicy -ExecutionPolicy Restricted -Force
}
catch {
    Write-Host "Error changing execution policy: $($_.Exception.Message)"
}
Stop-Transcript