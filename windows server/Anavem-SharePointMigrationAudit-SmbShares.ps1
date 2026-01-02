<#
================================================================================
 Anavem.com - PowerShell Script
================================================================================
 Script Name : Anavem-SharePointMigrationAudit-SmbShares.ps1
 Description : Audits SMB shares to prepare a migration to SharePoint Online.
               Exports CSV reports + optional HTML summary.

 Author      : Anavem.com
 Version     : 1.0.0
 Website     : https://www.anavem.com

 What it collects:
 - SMB share list (name, path, description, type, hidden)
 - Share permissions (Get-SmbShareAccess when available)
 - NTFS permissions (root, optional recursion to MaxDepth)
 - SharePoint compatibility checks (path length, invalid chars, blocked extensions, reserved names)
 - Content statistics (files/folders/size, top extensions, largest file, oldest/newest)

 Output:
 - CSV exports + HTML summary (optional) to the chosen OutputDir

 Notes:
 - For remote servers, local share paths are converted to UNC admin paths (e.g., \\SERVER\D$\Data).
   This requires administrative access and admin shares available.
 - Group resolution requires RSAT ActiveDirectory module and domain connectivity.
================================================================================
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$ServerName = $env:COMPUTERNAME,

    [Parameter()]
    [string]$OutputDir = "C:\Scripts",

    [Parameter()]
    [switch]$IncludeHiddenShares,

    [Parameter()]
    [ValidateRange(0, 50)]
    [int]$MaxDepth = 5,

    [Parameter()]
    [switch]$ResolveGroups,

    [Parameter()]
    [switch]$AnalyzePermissionsRecursively,

    [Parameter()]
    [switch]$GenerateHtmlReport,

    [Parameter()]
    [ValidateRange(0, 1000000000)]
    [int]$MaxItemsToScan = 0
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

# ----------------------------
# Paths / logging
# ----------------------------
if (-not (Test-Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}

$Timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$BaseName   = "Anavem-SharePointMigrationAudit_$ServerName" + "_$Timestamp"
$LogPath    = Join-Path $OutputDir "$BaseName.log"

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR")][string]$Level = "INFO"
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] [$Level] $Message"
    try { Add-Content -Path $LogPath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
    Write-Output $line
}

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
# SharePoint Online constraints (practical subset)
# ----------------------------
$Config = @{
    SharePoint = @{
        MaxPathLength     = 400
        MaxFileSizeBytes  = 250GB
        BlockedExtensions = @(
            ".exe",".dll",".bat",".cmd",".ps1",".vbs",".js",".jar",".msi",".aspx",".asmx",".ashx"
        )
        InvalidChars = @('"','*',':','<','>','?','/','\','|','#','%','~','&','{','}')
        ReservedNames = @(
            "CON","PRN","AUX","NUL",
            "COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9",
            "LPT1","LPT2","LPT3","LPT4","LPT5","LPT6","LPT7","LPT8","LPT9"
        )
    }
}

function Convert-Size {
    param([long]$Bytes)
    if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes Bytes"
}

function Convert-LocalPathToAdminUNC {
    param(
        [Parameter(Mandatory=$true)][string]$Server,
        [Parameter(Mandatory=$true)][string]$LocalPath
    )

    # If already UNC, return as-is
    if ($LocalPath -like "\\*") { return $LocalPath }

    # Match "C:\Folder\Sub"
    if ($LocalPath -match "^[A-Za-z]:\\") {
        $drive = $LocalPath.Substring(0,1).ToUpper()
        $rest  = $LocalPath.Substring(2)  # remove "C:"
        return "\\$Server\${drive}$${rest}"
    }

    # Unknown format, return original
    return $LocalPath
}

function Test-SharePointCompatibility {
    param(
        [Parameter(Mandatory=$true)][string]$FullPath,
        [Parameter(Mandatory=$true)][string]$Name,
        [long]$SizeBytes = 0
    )

    $issues = New-Object System.Collections.Generic.List[object]

    if ($FullPath.Length -gt $Config.SharePoint.MaxPathLength) {
        $issues.Add([pscustomobject]@{
            Type = "PathTooLong"
            Severity = "Error"
            Path = $FullPath
            Detail = "Path length is $($FullPath.Length) (max $($Config.SharePoint.MaxPathLength))"
            Recommendation = "Shorten the path or reorganize folders"
        })
    }

    if ($SizeBytes -gt $Config.SharePoint.MaxFileSizeBytes) {
        $issues.Add([pscustomobject]@{
            Type = "FileTooLarge"
            Severity = "Error"
            Path = $FullPath
            Detail = "File size is $(Convert-Size $SizeBytes) (max $(Convert-Size $Config.SharePoint.MaxFileSizeBytes))"
            Recommendation = "Split or compress the file"
        })
    }

    $ext = [System.IO.Path]::GetExtension($Name).ToLower()
    if ($ext -and ($Config.SharePoint.BlockedExtensions -contains $ext)) {
        $issues.Add([pscustomobject]@{
            Type = "BlockedExtension"
            Severity = "Warning"
            Path = $FullPath
            Detail = "Extension $ext is blocked by default"
            Recommendation = "Review policy or rename/convert the file"
        })
    }

    foreach ($c in $Config.SharePoint.InvalidChars) {
        if ($Name.Contains($c)) {
            $issues.Add([pscustomobject]@{
                Type = "InvalidCharacter"
                Severity = "Error"
                Path = $FullPath
                Detail = "Name contains invalid character: $c"
                Recommendation = "Rename the file/folder"
            })
            break
        }
    }

    $base = [System.IO.Path]::GetFileNameWithoutExtension($Name).ToUpper()
    if ($base -and ($Config.SharePoint.ReservedNames -contains $base)) {
        $issues.Add([pscustomobject]@{
            Type = "ReservedName"
            Severity = "Error"
            Path = $FullPath
            Detail = "Name is reserved: $base"
            Recommendation = "Rename the file/folder"
        })
    }

    if ($Name.StartsWith(" ") -or $Name.EndsWith(" ") -or $Name.StartsWith(".") -or $Name.EndsWith(".")) {
        $issues.Add([pscustomobject]@{
            Type = "InvalidNameFormat"
            Severity = "Error"
            Path = $FullPath
            Detail = "Name starts/ends with space or dot"
            Recommendation = "Rename the file/folder"
        })
    }

    return $issues.ToArray()
}

function Get-FileShares {
    param([Parameter(Mandatory=$true)][string]$Server)

    Write-Log "Collecting SMB shares from $Server..." "INFO"

    $shares = New-Object System.Collections.Generic.List[object]

    try {
        $hasGetSmbShare = [bool](Get-Command Get-SmbShare -ErrorAction SilentlyContinue)

        if ($hasGetSmbShare) {
            $smbShares = if ($Server -ieq $env:COMPUTERNAME) {
                Get-SmbShare
            } else {
                Invoke-Command -ComputerName $Server -ScriptBlock { Get-SmbShare }
            }

            foreach ($s in @($smbShares)) {
                if (-not $IncludeHiddenShares -and $s.Name.EndsWith('$')) { continue }
                if ($s.Name -in @("IPC$","print$")) { continue }

                $shares.Add([pscustomobject]@{
                    ServerName   = $Server
                    ShareName    = $s.Name
                    LocalPath    = $s.Path
                    AccessPath   = (Convert-LocalPathToAdminUNC -Server $Server -LocalPath $s.Path)
                    Description  = $s.Description
                    ShareType    = $s.ShareType
                    IsHidden     = [bool]$s.Name.EndsWith('$')
                    CurrentUsers = $s.CurrentUsers
                })
            }
        } else {
            # Fallback: WMI shares
            $wmiShares = Get-WmiObject -Class Win32_Share -ComputerName $Server -ErrorAction Stop
            foreach ($s in @($wmiShares)) {
                if (-not $IncludeHiddenShares -and $s.Name.EndsWith('$')) { continue }
                if ($s.Name -in @("IPC$","print$")) { continue }

                $shares.Add([pscustomobject]@{
                    ServerName   = $Server
                    ShareName    = $s.Name
                    LocalPath    = $s.Path
                    AccessPath   = (Convert-LocalPathToAdminUNC -Server $Server -LocalPath $s.Path)
                    Description  = $s.Description
                    ShareType    = $s.Type
                    IsHidden     = [bool]$s.Name.EndsWith('$')
                    CurrentUsers = $null
                })
            }
        }

        Write-Log "Shares found: $($shares.Count)" "INFO"
    }
    catch {
        Write-Log "Failed to collect shares: $($_.Exception.Message)" "ERROR"
    }

    return $shares.ToArray()
}

function Get-SharePermissions {
    param(
        [Parameter(Mandatory=$true)][string]$Server,
        [Parameter(Mandatory=$true)][string]$ShareName
    )

    $rows = New-Object System.Collections.Generic.List[object]
    try {
        if (-not (Get-Command Get-SmbShareAccess -ErrorAction SilentlyContinue)) {
            return $rows.ToArray()
        }

        $access = if ($Server -ieq $env:COMPUTERNAME) {
            Get-SmbShareAccess -Name $ShareName -ErrorAction SilentlyContinue
        } else {
            Invoke-Command -ComputerName $Server -ScriptBlock {
                param($n)
                Get-SmbShareAccess -Name $n -ErrorAction SilentlyContinue
            } -ArgumentList $ShareName
        }

        foreach ($ace in @($access)) {
            $rows.Add([pscustomobject]@{
                ShareName        = $ShareName
                AccountName      = $ace.AccountName
                AccessControlType= $ace.AccessControlType
                AccessRight      = $ace.AccessRight
            })
        }
    }
    catch {
        Write-Log "Share permissions failed for $ShareName: $($_.Exception.Message)" "WARN"
    }

    return $rows.ToArray()
}

function Get-NTFSPermissions {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$ShareName,
        [int]$Depth = 0
    )

    $rows = New-Object System.Collections.Generic.List[object]
    if ($MaxDepth -gt 0 -and $Depth -gt $MaxDepth) { return $rows.ToArray() }

    try {
        $acl = Get-Acl -Path $Path -ErrorAction Stop
        foreach ($ace in $acl.Access) {
            $rows.Add([pscustomobject]@{
                ShareName        = $ShareName
                Path             = $Path
                Identity         = $ace.IdentityReference.ToString()
                Rights           = $ace.FileSystemRights.ToString()
                Type             = $ace.AccessControlType.ToString()
                IsInherited      = [bool]$ace.IsInherited
                InheritanceFlags = $ace.InheritanceFlags.ToString()
                PropagationFlags = $ace.PropagationFlags.ToString()
                Depth            = $Depth
            })
        }

        if ($AnalyzePermissionsRecursively -and ($MaxDepth -gt 0) -and ($Depth -lt $MaxDepth)) {
            $subdirs = Get-ChildItem -Path $Path -Directory -Force -ErrorAction SilentlyContinue
            foreach ($d in @($subdirs)) {
                $child = Get-NTFSPermissions -Path $d.FullName -ShareName $ShareName -Depth ($Depth + 1)
                foreach ($x in @($child)) { $rows.Add($x) }
            }
        }
    }
    catch {
        Write-Log "NTFS permissions failed for $Path: $($_.Exception.Message)" "WARN"
    }

    return $rows.ToArray()
}

function Get-ShareStatistics {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$ShareName,
        [ref]$CompatibilityIssuesOut
    )

    Write-Log "Collecting statistics for share $ShareName..." "INFO"

    $totalSize = 0L
    $files = 0
    $folders = 0
    $oldest = $null
    $newest = $null
    $largestPath = $null
    $largestSize = 0L
    $extTable = @{}
    $issuesCount = 0
    $scanned = 0

    try {
        $items = Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue

        foreach ($item in @($items)) {
            if ($MaxItemsToScan -gt 0 -and $scanned -ge $MaxItemsToScan) { break }
            $scanned++

            $size = 0L
            if (-not $item.PSIsContainer) { $size = [long]$item.Length }

            $compatIssues = Test-SharePointCompatibility -FullPath $item.FullName -Name $item.Name -SizeBytes $size
            if ($compatIssues -and $compatIssues.Count -gt 0) {
                foreach ($ci in $compatIssues) { $CompatibilityIssuesOut.Value.Add($ci) }
                $issuesCount += $compatIssues.Count
            }

            if ($item.PSIsContainer) {
                $folders++
                continue
            }

            $files++
            $totalSize += $size

            $ext = ($item.Extension ?? "").ToLower()
            if (-not $extTable.ContainsKey($ext)) { $extTable[$ext] = @{ Count = 0; Size = 0L } }
            $extTable[$ext].Count++
            $extTable[$ext].Size += $size

            if ($size -gt $largestSize) { $largestSize = $size; $largestPath = $item.FullName }

            if ($null -eq $oldest -or $item.LastWriteTime -lt $oldest) { $oldest = $item.LastWriteTime }
            if ($null -eq $newest -or $item.LastWriteTime -gt $newest) { $newest = $item.LastWriteTime }
        }
    }
    catch {
        Write-Log "Statistics failed for $Path: $($_.Exception.Message)" "WARN"
    }

    $topExt = $extTable.GetEnumerator() |
        Sort-Object { $_.Value.Count } -Descending |
        Select-Object -First 10 |
        ForEach-Object {
            [pscustomobject]@{
                Extension = $_.Key
                Count     = $_.Value.Count
                SizeBytes = $_.Value.Size
                Size      = Convert-Size $_.Value.Size
            }
        }

    return [pscustomobject]@{
        ShareName           = $ShareName
        Path                = $Path
        TotalSizeBytes      = $totalSize
        TotalSize           = Convert-Size $totalSize
        TotalFiles          = $files
        TotalFolders        = $folders
        OldestWriteTime     = $oldest
        NewestWriteTime     = $newest
        LargestFilePath     = $largestPath
        LargestFileSize     = Convert-Size $largestSize
        TopExtensions       = ($topExt | ConvertTo-Json -Compress)
        CompatibilityIssues = $issuesCount
        ItemsScanned        = $scanned
    }
}

function Ensure-ADModuleIfNeeded {
    if (-not $ResolveGroups) { return $true }

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Log "ResolveGroups enabled but ActiveDirectory module not found. Group resolution will be skipped." "WARN"
        return $false
    }

    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        return $true
    } catch {
        Write-Log "Failed to import ActiveDirectory module. Group resolution will be skipped." "WARN"
        return $false
    }
}

function Get-ADGroupMembersRecursive {
    param(
        [Parameter(Mandatory=$true)][string]$GroupSam,
        [int]$Depth = 0,
        [int]$MaxDepthGroups = 3
    )

    if ($Depth -ge $MaxDepthGroups) { return @() }

    $members = New-Object System.Collections.Generic.List[object]

    try {
        $groupMembers = Get-ADGroupMember -Identity $GroupSam -ErrorAction Stop
        foreach ($m in @($groupMembers)) {
            if ($m.objectClass -eq "group") {
                $members.Add([pscustomobject]@{
                    ParentGroup    = $GroupSam
                    MemberName     = $m.Name
                    MemberType     = "Group"
                    SamAccountName = $m.SamAccountName
                    Depth          = $Depth
                })
                $child = Get-ADGroupMembersRecursive -GroupSam $m.SamAccountName -Depth ($Depth + 1) -MaxDepthGroups $MaxDepthGroups
                foreach ($c in @($child)) { $members.Add($c) }
            } else {
                $members.Add([pscustomobject]@{
                    ParentGroup    = $GroupSam
                    MemberName     = $m.Name
                    MemberType     = $m.objectClass
                    SamAccountName = $m.SamAccountName
                    Depth          = $Depth
                })
            }
        }
    } catch {
        Write-Log "Failed to resolve group $GroupSam: $($_.Exception.Message)" "WARN"
    }

    return $members.ToArray()
}

function Export-HtmlReportSimple {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][object[]]$Shares,
        [Parameter(Mandatory=$true)][object[]]$Stats,
        [Parameter(Mandatory=$true)][object[]]$Issues
    )

    $totalShares = $Shares.Count
    $totalSize   = ($Stats | Measure-Object -Property TotalSizeBytes -Sum).Sum
    $totalFiles  = ($Stats | Measure-Object -Property TotalFiles -Sum).Sum
    $totalIssues = $Issues.Count

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>SharePoint Migration Audit Report</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;background:#f6f7f9;color:#111;margin:0;padding:0}
.container{max-width:1200px;margin:0 auto;padding:24px}
.header{background:#0f172a;color:#fff;border-radius:10px;padding:18px 20px;margin-bottom:16px}
.h1{font-size:20px;margin:0 0 6px 0}
.meta{opacity:.9;font-size:13px}
.grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin:14px 0 18px}
.card{background:#fff;border-radius:10px;padding:12px 14px;border:1px solid #e6e8ee}
.k{font-size:12px;color:#555;text-transform:uppercase;letter-spacing:.03em}
.v{font-size:18px;font-weight:700;margin-top:6px}
table{width:100%;border-collapse:collapse;background:#fff;border:1px solid #e6e8ee;border-radius:10px;overflow:hidden}
th,td{padding:10px 12px;border-bottom:1px solid #eef1f6;font-size:13px;vertical-align:top}
th{background:#f2f4f8;text-align:left}
.badge{display:inline-block;padding:2px 10px;border-radius:999px;font-size:12px;border:1px solid #e6e8ee;background:#fff}
.badge.warn{border-color:#f59e0b}
.badge.err{border-color:#ef4444}
.section{margin:16px 0}
.section h2{font-size:16px;margin:0 0 10px 0}
.small{font-size:12px;color:#555}
</style>
</head>
<body>
<div class="container">
  <div class="header">
    <div class="h1">SharePoint Migration Audit Report</div>
    <div class="meta">Server: $ServerName | Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm") | Source: anavem.com</div>
  </div>

  <div class="grid">
    <div class="card"><div class="k">Shares</div><div class="v">$totalShares</div></div>
    <div class="card"><div class="k">Total size</div><div class="v">$(Convert-Size $totalSize)</div></div>
    <div class="card"><div class="k">Files</div><div class="v">$($totalFiles.ToString("N0"))</div></div>
    <div class="card"><div class="k">Compatibility issues</div><div class="v">$totalIssues</div></div>
  </div>

  <div class="section">
    <h2>Shares</h2>
    <table>
      <thead><tr>
        <th>Name</th><th>Local path</th><th>Access path</th><th>Description</th>
      </tr></thead>
      <tbody>
"@

    foreach ($s in $Shares) {
        $html += "<tr><td><b>$($s.ShareName)</b></td><td>$($s.LocalPath)</td><td>$($s.AccessPath)</td><td>$($s.Description)</td></tr>`n"
    }

    $html += @"
      </tbody>
    </table>
    <div class="small" style="margin-top:8px;">
      Access path uses administrative shares (e.g., \\SERVER\D$\Folder). Requires admin permissions.
    </div>
  </div>

  <div class="section">
    <h2>Top compatibility issues (first 200)</h2>
    <table>
      <thead><tr>
        <th>Severity</th><th>Type</th><th>Detail</th><th>Path</th><th>Recommendation</th>
      </tr></thead>
      <tbody>
"@

    foreach ($i in ($Issues | Select-Object -First 200)) {
        $cls = if ($i.Severity -eq "Error") { "err" } else { "warn" }
        $html += "<tr><td><span class='badge $cls'>$($i.Severity)</span></td><td>$($i.Type)</td><td>$($i.Detail)</td><td>$($i.Path)</td><td>$($i.Recommendation)</td></tr>`n"
    }

    $html += @"
      </tbody>
    </table>
  </div>

</div>
</body>
</html>
"@

    $html | Out-File -FilePath $Path -Encoding UTF8 -Force
}

# ----------------------------
# Main
# ----------------------------
Write-Log "Starting audit. Server=$ServerName OutputDir=$OutputDir" "INFO"

# Connectivity check
try {
    if (-not (Test-Connection -ComputerName $ServerName -Count 1 -Quiet)) {
        Write-Log "Server is not reachable: $ServerName" "ERROR"
        exit 1
    }
} catch {
    Write-Log "Connectivity test failed: $($_.Exception.Message)" "ERROR"
    exit 1
}

# Collect shares
$shares = Get-FileShares -Server $ServerName
if (-not $shares -or $shares.Count -eq 0) {
    Write-Log "No shares found." "WARN"
    exit 0
}

$sharePerms = New-Object System.Collections.Generic.List[object]
$ntfsPerms  = New-Object System.Collections.Generic.List[object]
$stats      = New-Object System.Collections.Generic.List[object]
$issues     = New-Object System.Collections.Generic.List[object]
$groupMembers = New-Object System.Collections.Generic.List[object]

$adOk = Ensure-ADModuleIfNeeded

$idx = 0
foreach ($sh in $shares) {
    $idx++
    Write-Log "Processing share $idx/$($shares.Count): $($sh.ShareName)" "INFO"

    # Share permissions
    $sp = Get-SharePermissions -Server $ServerName -ShareName $sh.ShareName
    foreach ($x in @($sp)) { $sharePerms.Add($x) }

    # NTFS + stats require accessible path
    $accessPath = $sh.AccessPath
    if (-not (Test-Path $accessPath)) {
        Write-Log "Path not accessible: $accessPath" "WARN"
        continue
    }

    # NTFS permissions
    $np = Get-NTFSPermissions -Path $accessPath -ShareName $sh.ShareName -Depth 0
    foreach ($x in @($np)) { $ntfsPerms.Add($x) }

    # Statistics + compatibility issues
    $st = Get-ShareStatistics -Path $accessPath -ShareName $sh.ShareName -CompatibilityIssuesOut ([ref]$issues)
    $stats.Add($st)
}

# Optional: group resolution (based on permissions)
if ($ResolveGroups -and $adOk) {
    Write-Log "Resolving AD group membership (best effort)..." "INFO"

    $candidates = @()

    $candidates += ($ntfsPerms | Select-Object -ExpandProperty Identity -ErrorAction SilentlyContinue)
    $candidates += ($sharePerms | Select-Object -ExpandProperty AccountName -ErrorAction SilentlyContinue)

    $unique = $candidates |
        Where-Object { $_ -and $_ -notmatch "^(NT AUTHORITY|BUILTIN|S-1-)" } |
        ForEach-Object { ($_ -replace "^.*\\","").Trim() } |
        Select-Object -Unique

    foreach ($g in $unique) {
        $members = Get-ADGroupMembersRecursive -GroupSam $g
        foreach ($m in @($members)) { $groupMembers.Add($m) }
    }
}

# ----------------------------
# Exports
# ----------------------------
Write-Log "Exporting results..." "INFO"

$csvShares     = Join-Path $OutputDir "$BaseName_Shares.csv"
$csvSharePerms = Join-Path $OutputDir "$BaseName_SharePermissions.csv"
$csvNtfsPerms  = Join-Path $OutputDir "$BaseName_NTFSPermissions.csv"
$csvStats      = Join-Path $OutputDir "$BaseName_Statistics.csv"
$csvIssues     = Join-Path $OutputDir "$BaseName_CompatibilityIssues.csv"
$csvGroups     = Join-Path $OutputDir "$BaseName_GroupMembers.csv"
$htmlReport    = Join-Path $OutputDir "$BaseName_Report.html"

Export-CsvAlways -Path $csvShares -Data $shares -HeaderObject @{
    ServerName=""; ShareName=""; LocalPath=""; AccessPath=""; Description=""; ShareType=""; IsHidden=$false; CurrentUsers=""
}

Export-CsvAlways -Path $csvSharePerms -Data $sharePerms.ToArray() -HeaderObject @{
    ShareName=""; AccountName=""; AccessControlType=""; AccessRight=""
}

Export-CsvAlways -Path $csvNtfsPerms -Data $ntfsPerms.ToArray() -HeaderObject @{
    ShareName=""; Path=""; Identity=""; Rights=""; Type=""; IsInherited=$false; InheritanceFlags=""; PropagationFlags=""; Depth=0
}

Export-CsvAlways -Path $csvStats -Data $stats.ToArray() -HeaderObject @{
    ShareName=""; Path=""; TotalSizeBytes=0; TotalSize=""; TotalFiles=0; TotalFolders=0; OldestWriteTime=""; NewestWriteTime="";
    LargestFilePath=""; LargestFileSize=""; TopExtensions=""; CompatibilityIssues=0; ItemsScanned=0
}

Export-CsvAlways -Path $csvIssues -Data $issues.ToArray() -HeaderObject @{
    Type=""; Severity=""; Path=""; Detail=""; Recommendation=""
}

if ($groupMembers.Count -gt 0) {
    Export-CsvAlways -Path $csvGroups -Data $groupMembers.ToArray() -HeaderObject @{
        ParentGroup=""; MemberName=""; MemberType=""; SamAccountName=""; Depth=0
    }
}

if ($GenerateHtmlReport) {
    try {
        Export-HtmlReportSimple -Path $htmlReport -Shares $shares -Stats $stats.ToArray() -Issues $issues.ToArray()
        Write-Log "HTML report created: $htmlReport" "INFO"
    } catch {
        Write-Log "HTML report generation failed: $($_.Exception.Message)" "WARN"
    }
}

Write-Log "Done." "INFO"
Write-Log "Outputs:" "INFO"
Write-Log " - $csvShares" "INFO"
Write-Log " - $csvSharePerms" "INFO"
Write-Log " - $csvNtfsPerms" "INFO"
Write-Log " - $csvStats" "INFO"
Write-Log " - $csvIssues" "INFO"
if ($groupMembers.Count -gt 0) { Write-Log " - $csvGroups" "INFO" }
if ($GenerateHtmlReport -and (Test-Path $htmlReport)) { Write-Log " - $htmlReport" "INFO" }
Write-Log "Log: $LogPath" "INFO"
