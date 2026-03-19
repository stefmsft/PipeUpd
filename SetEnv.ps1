# SetEnv.ps1 - Environment Switcher
# Lists available .env files and writes selection to config.ini

$envFiles = Get-ChildItem -Path "." -Filter ".env*" -File | Where-Object { $_.Name -ne ".env.template" } | Sort-Object Name

if ($envFiles.Count -eq 0) {
    Write-Host "No .env files found in current directory." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  SetEnv - Environment Switcher" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
# Read current value from config.ini
$currentEnv = $null
if (Test-Path "config.ini") {
    $line = Get-Content "config.ini" | Where-Object { $_ -match "^\s*ENV_SUFFIX\s*=" } | Select-Object -First 1
    if ($line) {
        $currentSuffix = ($line -split "=", 2)[1].Trim()
        if ($currentSuffix -eq "") {
            $currentEnv = "Production  (.env)"
        } else {
            $currentEnv = "$currentSuffix  (.env.$currentSuffix)"
        }
    }
}

Write-Host ""
if ($currentEnv) {
    Write-Host "Current environment: $currentEnv" -ForegroundColor Magenta
} else {
    Write-Host "Current environment: (not set)" -ForegroundColor DarkGray
}
Write-Host ""
Write-Host "Available environments:" -ForegroundColor Yellow
Write-Host ""

$menuItems = @()
for ($i = 0; $i -lt $envFiles.Count; $i++) {
    $file = $envFiles[$i]
    if ($file.Name -eq ".env") {
        $displayName = "Production"
        $suffix = ""
    }
    else {
        $suffix = $file.Name.Substring(5)  # Remove ".env." prefix
        $displayName = $suffix
    }
    $menuItems += [PSCustomObject]@{
        Index       = $i + 1
        DisplayName = $displayName
        Suffix      = $suffix
        FileName    = $file.Name
    }
    Write-Host "  $($i + 1). $displayName  ($($file.Name))" -ForegroundColor White
}

Write-Host ""
$choice = Read-Host "Select environment (1-$($envFiles.Count))"

# Validate input
$choiceNum = 0
if ($choice -eq "") {
    Write-Host "No selection made. Exiting without changes." -ForegroundColor DarkGray
    exit 0
}
if (-not [int]::TryParse($choice, [ref]$choiceNum) -or $choiceNum -lt 1 -or $choiceNum -gt $envFiles.Count) {
    Write-Host "Invalid selection. No changes made." -ForegroundColor Red
    exit 1
}

$selected = $menuItems[$choiceNum - 1]

# Write config.ini
$configContent = @"
[Environment]
ENV_SUFFIX=$($selected.Suffix)
"@

Set-Content -Path "config.ini" -Value $configContent -Encoding UTF8

Write-Host ""
Write-Host "Environment set to: $($selected.DisplayName)  ($($selected.FileName))" -ForegroundColor Green
Write-Host "Written to config.ini" -ForegroundColor Green
Write-Host ""
