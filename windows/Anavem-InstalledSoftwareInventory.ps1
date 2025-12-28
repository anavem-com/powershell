<# 
================================================================================
 Anavem.com - PowerShell Script
================================================================================
 Script Name : Anavem-InstalledSoftwareInventory.ps1
 Description : Exports installed software inventory (32-bit & 64-bit) to TXT
 Author      : Anavem.com
 Version     : 1.0.0
 Website     : https://www.anavem.com

 Use Case    :
 - Software inventory
 - License audits
 - IT support diagnostics
 - Asset documentation
 - MSP / RMM reporting

 Run as Admin: No
================================================================================
#>

# Ensure output directory exists
$OutputDir = "C:\Scripts"
if (-not (Test-Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}

# Output file
$Timestamp  = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$OutputFile = "$OutputDir\Anavem-InstalledSoftwareInventory_$Timestamp.txt"

# Collect installed software (32-bit & 64-bit)
$Software = Get-ItemProperty `
    HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,
    HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
    Where-Object { $_.DisplayName } |
    Select-Object DisplayName, DisplayVersion, Publisher, InstallDate |
    Sort-Object DisplayName

# Header
$Header = @"
Anavem.com : Installed Software Inventory
Script   : Anavem-InstalledSoftwareInventory.ps1
Version  : 1.0.0
Generated: $(Get-Date)
------------------------------------------------------------

"@

# Write output
$Header | Out-File -FilePath $OutputFile -Encoding UTF8

$Software |
Format-Table DisplayName, DisplayVersion, Publisher, InstallDate -AutoSize |
Out-String |
Out-File -FilePath $OutputFile -Append -Encoding UTF8

Write-Host "Installed software inventory saved to $OutputFile"
