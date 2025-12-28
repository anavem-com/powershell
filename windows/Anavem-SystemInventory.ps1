<# 
================================================================================
 Anavem.com - PowerShell Script
================================================================================
 Script Name : Anavem-SystemInventory.ps1
 Description : Collects core Windows system and hardware information
 Author      : Anavem.com
 Version     : 1.0.0
 Last Update : 2025-01-28
 Website     : https://www.anavem.com
 GitHub      : https://github.com/anavem-com

 Use Case    :
 - Quick system inventory
 - IT support diagnostics
 - Intune / RMM reporting
 - Asset documentation

 Requirements:
 - Windows 10 / 11 / Windows Server
 - PowerShell 5.1+

 Run as Admin: No
================================================================================
#>

Write-Host "Anavem.com - System Inventory Script" -ForegroundColor Cyan
Write-Host "Script : Anavem-SystemInventory.ps1" -ForegroundColor Cyan
Write-Host "Version: 1.0.0" -ForegroundColor Cyan
Write-Host "Purpose: Display hardware and OS inventory information" -ForegroundColor Cyan
Write-Host "---------------------------------------------" -ForegroundColor DarkGray

$cs   = Get-CimInstance Win32_ComputerSystem
$os   = Get-CimInstance Win32_OperatingSystem
$cpu  = Get-CimInstance Win32_Processor | Select-Object -First 1
$bios = Get-CimInstance Win32_BIOS

[pscustomobject]@{
    ComputerName = $env:COMPUTERNAME
    Manufacturer = $cs.Manufacturer
    Model        = $cs.Model
    OperatingSystem = $os.Caption
    OSVersion    = $os.Version
    OSBuild      = $os.BuildNumber
    BIOSVersion  = $bios.SMBIOSBIOSVersion
    SerialNumber = $bios.SerialNumber
    CPU          = $cpu.Name
    RAM_GB       = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
    Uptime       = (Get-Date) - $os.LastBootUpTime
} | Format-List
