<# 
================================================================================
 Anavem.com - PowerShell Script
================================================================================
 Script Name : Anavem-Win11-Debloat.ps1
 Description : Windows 11 debloat script to remove selected built-in apps.
               Exports before/after report to C:\Scripts (CSV + TXT + LOG).
 Author      : Anavem.com
 Version     : 1.0.4
 Website     : https://www.anavem.com

 Requirements:
 - Windows 11
 - PowerShell 5.1+
 - Admin recommended for -Scope AllUsers and -Deprovision
================================================================================
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [ValidateSet("CurrentUser","AllUsers")]
    [string]$Scope = "CurrentUser",

    [switch]$Deprovision,
    [switch]$RemoveOneDrive,
    [string]$OutDir = "C:\Scripts"
)

# Keep execution resilient like your attached scripts
Set-StrictMode -Version 2.0
$global:ErrorActionPreference = "Continue"

# ----------------------------
# Output and logging
# ----------------------------
if (-not (Test-Path $OutDir)) {
    New-Item -Path $OutDir -ItemType Directory -Force | Out-Null
}

$Timestamp  = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$BaseName   = "Anavem-Win11-Debloat"
$LogFile    = Join-Path $OutDir "$BaseName.log"
$ReportFile = Join-Path $OutDir "$BaseName-Report_$Timestamp.txt"
$CsvBefore  = Join-Path $OutDir "$BaseName-Before_$Timestamp.csv"
$CsvAfter   = Join-Path $OutDir "$BaseName-After_$Timestamp.csv"
$CsvRemoved = Join-Path $OutDir "$BaseName-Removed_$Timestamp.csv"
$CsvErrors  = Join-Path $OutDir "$BaseName-Errors_$Timestamp.csv"

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet("INFO","WARNING","ERROR","DEBUG")][string]$Level="INFO"
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] [$Level] $Message"
    try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

function Get-OsContext {
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue

    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
               ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    [pscustomobject]@{
        ComputerName   = $env:COMPUTERNAME
        OS             = $os.Caption
        Version        = $os.Version
        Build          = $os.BuildNumber
        IsAdmin        = $isAdmin
        Scope          = $Scope
        Deprovision    = [bool]$Deprovision
        RemoveOneDrive = [bool]$RemoveOneDrive
        OutDir         = $OutDir
        PartOfDomain   = [bool]$cs.PartOfDomain
        Domain         = [string]$cs.Domain
    }
}

# ----------------------------
# Targets (excluding Store, Xbox.TCUI, Edge)
# ----------------------------
$Targets_Misc = @(
    "Microsoft.GetHelp",
    "Microsoft.M365Companions",
    "Microsoft.MSPaint",
    "Microsoft.OutlookForWindows",
    "Microsoft.OneDrive",
    "Microsoft.Paint",
    "Microsoft.People",
    "Microsoft.RemoteDesktop",
    "Microsoft.ScreenSketch",
    "Microsoft.Whiteboard",
    "Microsoft.Windows.Photos",
    "Microsoft.WindowsCalculator",
    "Microsoft.WindowsCamera",
    "Microsoft.WindowsNotepad",
    "Microsoft.windowscommunicationsapps",
    "Microsoft.WindowsTerminal",
    "Microsoft.YourPhone",
    "Microsoft.ZuneMusic",
    "MicrosoftWindows.CrossDevice"
)

$Targets_Gaming = @(
    "Microsoft.GamingApp",
    "Microsoft.XboxGameOverlay",
    "Microsoft.XboxGamingOverlay",
    "Microsoft.XboxIdentityProvider",
    "Microsoft.XboxSpeechToTextOverlay"
)

$Targets = $Targets_Misc + $Targets_Gaming

# ----------------------------
# CSV helpers (always create files)
# ----------------------------
function Export-CsvAlways {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][object[]]$Data,
        [Parameter(Mandatory=$true)][hashtable]$HeaderObject
    )

    if (-not $Data -or $Data.Count -eq 0) {
        [pscustomobject]$HeaderObject | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Force
        return
    }

    $Data | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Force
}

# ----------------------------
# Snapshot + removal
# ----------------------------
function Get-AppxSnapshot {
    param([string[]]$Names)

    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($n in $Names) {
        try {
            $pkgs = @()
            if ($Scope -eq "AllUsers") {
                $pkgs = Get-AppxPackage -AllUsers -Name $n -ErrorAction SilentlyContinue
            } else {
                $pkgs = Get-AppxPackage -Name $n -ErrorAction SilentlyContinue
            }

            foreach ($p in @($pkgs)) {
                $rows.Add([pscustomobject]@{
                    Name            = $n
                    PackageFullName = $p.PackageFullName
                    Version         = [string]$p.Version
                    Publisher       = [string]$p.Publisher
                    InstallLocation = [string]$p.InstallLocation
                })
            }
        } catch {
            Write-Log "Snapshot error for $n : $($_.Exception.Message)" "WARNING"
        }
    }

    return $rows.ToArray()
}

function Remove-AppxByName {
    param([Parameter(Mandatory=$true)][string]$Name)

    $results = New-Object System.Collections.Generic.List[object]

    $pkgs = @()
    try {
        if ($Scope -eq "AllUsers") {
            $pkgs = Get-AppxPackage -AllUsers -Name $Name -ErrorAction SilentlyContinue
        } else {
            $pkgs = Get-AppxPackage -Name $Name -ErrorAction SilentlyContinue
        }
    } catch {
        return $results.ToArray()
    }

    foreach ($p in @($pkgs)) {
        $id = $p.PackageFullName
        try {
            if ($PSCmdlet.ShouldProcess($id, "Remove-AppxPackage")) {
                Remove-AppxPackage -Package $id -ErrorAction SilentlyContinue
                Write-Log "Removed Appx package: $id" "INFO"
                $results.Add([pscustomobject]@{
                    Name            = $Name
                    PackageFullName = $id
                    Action          = "Remove-AppxPackage"
                    Result          = "Attempted"
                    Error           = ""
                })
            }
        } catch {
            $results.Add([pscustomobject]@{
                Name            = $Name
                PackageFullName = $id
                Action          = "Remove-AppxPackage"
                Result          = "Failed"
                Error           = $_.Exception.Message
            })
        }
    }

    if ($Deprovision) {
        try {
            $prov = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
                    Where-Object { $_.DisplayName -eq $Name }

            foreach ($pr in @($prov)) {
                if ($PSCmdlet.ShouldProcess($pr.PackageName, "Remove-AppxProvisionedPackage")) {
                    Remove-AppxProvisionedPackage -Online -PackageName $pr.PackageName -ErrorAction SilentlyContinue | Out-Null
                    Write-Log "Deprovisioned package: $($pr.PackageName) (DisplayName=$Name)" "INFO"
                    $results.Add([pscustomobject]@{
                        Name            = $Name
                        PackageFullName = $pr.PackageName
                        Action          = "Remove-AppxProvisionedPackage"
                        Result          = "Attempted"
                        Error           = ""
                    })
                }
            }
        } catch {
            $results.Add([pscustomobject]@{
                Name            = $Name
                PackageFullName = ""
                Action          = "Remove-AppxProvisionedPackage"
                Result          = "Failed"
                Error           = $_.Exception.Message
            })
        }
    }

    return $results.ToArray()
}

function Uninstall-OneDrive {
    $candidates = @(
        "$env:SystemRoot\System32\OneDriveSetup.exe",
        "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
    )
    $exe = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $exe) {
        Write-Log "OneDriveSetup.exe not found. Skipping OneDrive uninstall." "WARNING"
        return [pscustomobject]@{
            Name            = "Microsoft.OneDrive"
            PackageFullName = ""
            Action          = "OneDriveSetup.exe /uninstall"
            Result          = "Skipped"
            Error           = "OneDriveSetup.exe not found"
        }
    }

    try {
        if ($PSCmdlet.ShouldProcess($exe, "Uninstall OneDrive")) {
            Start-Process -FilePath $exe -ArgumentList "/uninstall" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
            Write-Log "OneDrive uninstall executed: $exe /uninstall" "INFO"
            return [pscustomobject]@{
                Name            = "Microsoft.OneDrive"
                PackageFullName = ""
                Action          = "OneDriveSetup.exe /uninstall"
                Result          = "Attempted"
                Error           = ""
            }
        }
    } catch {
        return [pscustomobject]@{
            Name            = "Microsoft.OneDrive"
            PackageFullName = ""
            Action          = "OneDriveSetup.exe /uninstall"
            Result          = "Failed"
            Error           = $_.Exception.Message
        }
    }
}

function Group-Counts {
    param([object[]]$rows)
    if (-not $rows -or $rows.Count -eq 0) { return @() }
    $rows | Group-Object Name | ForEach-Object {
        [pscustomobject]@{ Name=$_.Name; Count=$_.Count }
    } | Sort-Object Name
}

# ----------------------------
# Run (ensure AFTER CSV is always created)
# ----------------------------
$ctx = Get-OsContext
Write-Log "Starting $BaseName v1.0.4" "INFO"
Write-Log ("Context: " + ($ctx | ConvertTo-Json -Compress)) "INFO"

$before = @()
$after  = @()
$removedAll = New-Object System.Collections.Generic.List[object]
$errorsAll  = New-Object System.Collections.Generic.List[object]

try {
    # BEFORE snapshot
    $before = Get-AppxSnapshot -Names $Targets
    Export-CsvAlways -Path $CsvBefore -Data $before -HeaderObject @{
        Name=""; PackageFullName=""; Version=""; Publisher=""; InstallLocation=""
    }
    Write-Log "Before CSV created: $CsvBefore" "INFO"

    # Removals
    foreach ($t in $Targets) {
        try {
            $res = Remove-AppxByName -Name $t
            foreach ($r in @($res)) { $removedAll.Add($r) }
        } catch {
            $errorsAll.Add([pscustomobject]@{ Target=$t; Stage="Remove-AppxByName"; Error=$_.Exception.Message })
            Write-Log "Removal stage error for $t : $($_.Exception.Message)" "WARNING"
        }
    }

    if ($RemoveOneDrive) {
        try {
            $removedAll.Add((Uninstall-OneDrive))
        } catch {
            $errorsAll.Add([pscustomobject]@{ Target="Microsoft.OneDrive"; Stage="Uninstall-OneDrive"; Error=$_.Exception.Message })
        }
    }

} finally {
    # AFTER snapshot MUST happen even if removals fail
    $after = Get-AppxSnapshot -Names $Targets
    Export-CsvAlways -Path $CsvAfter -Data $after -HeaderObject @{
        Name=""; PackageFullName=""; Version=""; Publisher=""; InstallLocation=""
    }
    Write-Log "After CSV created: $CsvAfter" "INFO"

    # Removed CSV always
    Export-CsvAlways -Path $CsvRemoved -Data $removedAll.ToArray() -HeaderObject @{
        Name=""; PackageFullName=""; Action=""; Result=""; Error=""
    }
    Write-Log "Removed CSV created: $CsvRemoved" "INFO"

    # Errors CSV always
    Export-CsvAlways -Path $CsvErrors -Data $errorsAll.ToArray() -HeaderObject @{
        Target=""; Stage=""; Error=""
    }
    Write-Log "Errors CSV created: $CsvErrors" "INFO"

    # Report
    $beforeCounts = Group-Counts $before
    $afterCounts  = Group-Counts $after

    $lines = @()
    $lines += "Anavem.com - Windows 11 Debloat Report"
    $lines += "Script    : Anavem-Win11-Debloat.ps1"
    $lines += "Version   : 1.0.4"
    $lines += "Generated : $(Get-Date)"
    $lines += "------------------------------------------------------------"
    $lines += "Scope          : $Scope"
    $lines += "Deprovision    : $Deprovision"
    $lines += "RemoveOneDrive : $RemoveOneDrive"
    $lines += "Output Dir     : $OutDir"
    $lines += "------------------------------------------------------------"
    $lines += ""
    $lines += "FILES"
    $lines += "Before CSV : $CsvBefore"
    $lines += "After CSV  : $CsvAfter"
    $lines += "Removed CSV: $CsvRemoved"
    $lines += "Errors CSV : $CsvErrors"
    $lines += "Log File   : $LogFile"
    $lines += "------------------------------------------------------------"
    $lines += ""
    $lines += "TARGET COUNTS - BEFORE"
    if ($beforeCounts.Count -gt 0) { $lines += ($beforeCounts | Format-Table Name, Count -AutoSize | Out-String) }
    else { $lines += "No matching packages found in BEFORE snapshot." }

    $lines += "TARGET COUNTS - AFTER"
    if ($afterCounts.Count -gt 0) { $lines += ($afterCounts | Format-Table Name, Count -AutoSize | Out-String) }
    else { $lines += "No matching packages found in AFTER snapshot." }

    $lines += ""
    $lines += "NOTES"
    $lines += "- Some applications may be reinstalled by Microsoft 365 or enterprise policies."
    $lines += "- For golden images, consider using -Scope AllUsers with -Deprovision (admin recommended)."
    $lines += "- Validate results on a test device before large-scale deployment."

    $lines | Out-File -Path $ReportFile -Encoding UTF8 -Force
    Write-Log "Completed. Report: $ReportFile" "INFO"
}

Write-Output "Done. Report=$ReportFile BeforeCsv=$CsvBefore AfterCsv=$CsvAfter RemovedCsv=$CsvRemoved ErrorsCsv=$CsvErrors"
