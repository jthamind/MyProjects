<#
.SYNOPSIS
  Sync IIS role services (Web-*) from a source ("golden") server to this target server.
.DESCRIPTION
  1) Fetches installed IIS Windows Features (Name like "Web-*") from the source server.
  2) Compares to the locally installed IIS Windows Features.
  3) Installs any missing features (with management tools).
  4) Reports on any extra IIS modules (global or application) present on source but not on target.
#>

# — VARIABLES: change these to match your environment —
$sourceServer = 'SECOND_SERVER_NAME'
# (use hostname or FQDN that you can reach via WinRM)

# — ENSURE RUNNING AS ADMINISTRATOR —
$currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal       = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Warning "ERROR: This script must be run as Administrator!"
    exit 1
}

# — IMPORT MODULES —
Import-Module ServerManager       -ErrorAction Stop   # Get-WindowsFeature
Import-Module WebAdministration   -ErrorAction Stop   # IIS cmdlets

# — 1) GET IIS FEATURES —
Write-Host "Fetching IIS features from source server '$sourceServer'..." -ForegroundColor Cyan
$sourceFeatures = Invoke-Command -ComputerName $sourceServer -ScriptBlock {
    Import-Module ServerManager
    Get-WindowsFeature | Where-Object { $_.Name -like 'Web-*' -and $_.Installed } | Select-Object -ExpandProperty Name
}

Write-Host "Fetching IIS features from LOCAL server..." -ForegroundColor Cyan
$localFeatures = Get-WindowsFeature | Where-Object { $_.Name -like 'Web-*' -and $_.Installed } | Select-Object -ExpandProperty Name

# — 2) DETECT MISSING FEATURES —
$missingFeatures = $sourceFeatures | Where-Object { $_ -notin $localFeatures }

if ($missingFeatures.Count -gt 0) {
    Write-Host "Missing IIS Windows Features on THIS server:" -ForegroundColor Yellow
    $missingFeatures | ForEach-Object { Write-Host "  - $_" }
    Write-Host "`nInstalling missing features (and Management Tools)..." -ForegroundColor Cyan
    Install-WindowsFeature `
      -Name $missingFeatures `
      -IncludeManagementTools `
      -ErrorAction Stop |
      Format-Table -AutoSize

    Write-Host "`nRestarting IIS to pick up new modules..." -ForegroundColor Cyan
    iisreset
}
else {
    Write-Host "All IIS Windows Features are already in sync!" -ForegroundColor Green
}

# — 3) OPTIONAL: COMPARE GLOBAL MODULES —
Write-Host "`nChecking IIS Global Modules..." -ForegroundColor Cyan
$sourceGlobal = Invoke-Command -ComputerName $sourceServer -ScriptBlock {
    Import-Module WebAdministration
    Get-WebGlobalModule | Select-Object -ExpandProperty Name
}
$localGlobal = Get-WebGlobalModule | Select-Object -ExpandProperty Name

$missingGlobal = $sourceGlobal | Where-Object { $_ -notin $localGlobal }
if ($missingGlobal) {
    Write-Warning "The following Global Modules exist on source but not locally:"
    $missingGlobal | ForEach-Object { Write-Host "  - $_" }
    Write-Warning "Most of these are installed by the Windows Features above.  If any remain, you may need to install them manually or register via Add-WebGlobalModule."
} else {
    Write-Host "Global modules are in sync." -ForegroundColor Green
}

# — 4) OPTIONAL: COMPARE APPLICATION-LEVEL MODULES —
Write-Host "`nChecking IIS Application Modules (/system.webServer/modules)..." -ForegroundColor Cyan
$sourceApp = Invoke-Command -ComputerName $sourceServer -ScriptBlock {
    Import-Module WebAdministration
    Get-WebConfiguration -Filter "/system.webServer/modules/add" -PSPath "MACHINE/WEBROOT/APPHOST" |
      Select-Object -ExpandProperty attributes |
      ForEach-Object { $_['name'].Value }
}
$localApp = Get-WebConfiguration -Filter "/system.webServer/modules/add" -PSPath "MACHINE/WEBROOT/APPHOST" |
  Select-Object -ExpandProperty attributes |
  ForEach-Object { $_['name'].Value }

$missingApp = $sourceApp | Where-Object { $_ -notin $localApp }
if ($missingApp) {
    Write-Warning "The following application modules exist on source but not locally:"
    $missingApp | ForEach-Object { Write-Host "  - $_" }
    Write-Warning "These, too, typically come from the IIS role services above.  Custom modules (e.g. AppOptics, ServiceModel) must be installed separately."
} else {
    Write-Host "Application modules are in sync." -ForegroundColor Green
}

Write-Host "`n✅ Sync complete." -ForegroundColor Cyan
