<# 
================================================================================
 Anavem.com - PowerShell Script
================================================================================
 Script Name : Anavem-UpdateImmutableId.ps1
 Description : Updates the Immutable ID (onPremisesImmutableId) for a Microsoft
               365 user based on their Active Directory objectGUID. Enables 
               hard-matching between AD and Entra ID for directory sync.
 Author      : Anavem.com
 Version     : 1.0.0
 Website     : https://www.anavem.com

 Use Case    :
 - Fix Entra Connect sync errors (InvalidSoftMatch, attribute conflicts)
 - Enable hard-matching for cloud-only users before sync
 - Recover from AD forest migrations
 - Restore deleted users with correct sync identity

 Run as Admin: No (but requires User Administrator role in Entra ID)
================================================================================
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$UPN,

    [Parameter(Mandatory = $false)]
    [bool]$SkipConnection = $false,

    [Parameter(Mandatory = $false)]
    [bool]$Clear = $false
)

# Check for required modules
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    Write-Host "Microsoft.Graph.Users module not found." -ForegroundColor Red
    Write-Host "Install with: Install-Module Microsoft.Graph.Users -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "ActiveDirectory module not found." -ForegroundColor Red
    Write-Host "Install RSAT tools to get this module." -ForegroundColor Yellow
    exit 1
}

Import-Module ActiveDirectory -ErrorAction Stop

# Prompt for UPN if not provided
if ([string]::IsNullOrEmpty($UPN)) {
    $UPN = Read-Host "Enter User Principal Name (UPN)"
}

# Connect to Microsoft Graph
if (-not $SkipConnection) {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome -ErrorAction Stop
}

# Get AD user and objectGUID
$adUser = Get-ADUser -Filter "UserPrincipalName -eq '$UPN'" -Properties ObjectGUID

if ($null -eq $adUser) {
    Write-Host "User not found in Active Directory: $UPN" -ForegroundColor Red
    exit 1
}

# Convert objectGUID to ImmutableId (Base64)
$immutableId = [System.Convert]::ToBase64String($adUser.ObjectGUID.ToByteArray())
Write-Host "Generated ImmutableId: $immutableId" -ForegroundColor Green

# Get Microsoft 365 user
$mgUser = Get-MgUser -Filter "userPrincipalName eq '$UPN'" -Property Id, UserPrincipalName, OnPremisesImmutableId, OnPremisesSyncEnabled

if ($null -eq $mgUser) {
    Write-Host "User not found in Microsoft 365: $UPN" -ForegroundColor Red
    exit 1
}

# Check if user is currently synced
if ($mgUser.OnPremisesSyncEnabled -eq $true) {
    Write-Host "User is currently synced. Cannot update ImmutableId while sync is active." -ForegroundColor Red
    Write-Host "Disable sync for this user or move them out of sync scope first." -ForegroundColor Yellow
    exit 1
}

# Update or clear ImmutableId
if ($Clear) {
    Update-MgUser -UserId $mgUser.Id -OnPremisesImmutableId $null
    Write-Host "ImmutableId cleared for Microsoft 365 User." -ForegroundColor Green
}
else {
    Update-MgUser -UserId $mgUser.Id -OnPremisesImmutableId $immutableId
    Write-Host "ImmutableId written to Microsoft 365 User. Please confirm it matches generated ImmutableId" -ForegroundColor Green
}

# Verify and display result
$verifyUser = Get-MgUser -UserId $mgUser.Id -Property UserPrincipalName, OnPremisesImmutableId, OnPremisesSyncEnabled

$verifyUser | Select-Object UserPrincipalName, 
    @{N='ImmutableId';E={$_.OnPremisesImmutableId}}, 
    @{N='SyncEnabled';E={$_.OnPremisesSyncEnabled}} | Format-Table -AutoSize
