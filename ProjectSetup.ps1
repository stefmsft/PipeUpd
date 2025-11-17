#!/usr/bin/env pwsh

# Python Boilerplate Setup Script
# This script installs uv and sets up a Python development environment

# Parse command line arguments
param(
    [switch]$y,
    [switch]$help,
    [switch]$SkipPwshCheck  # Internal flag to skip PowerShell version check after reinstall
)

# Set strict mode to exit on any error
$ErrorActionPreference = "Stop"

# Show help if requested
if ($help) {
    Write-Host "Python Boilerplate Setup Script" -ForegroundColor Green
    Write-Host "Usage: .\ProjectSetup.ps1 [-y] [-help]" -ForegroundColor White
    Write-Host "  -y      : Non-interactive mode (use defaults)" -ForegroundColor White
    Write-Host "  -help   : Show this help message" -ForegroundColor White
    exit 0
}

# ============================================================================
# PowerShell Version Check and Installation
# ============================================================================
# This project requires PowerShell 7.5 or higher for proper functionality
# (Unicode support, better error handling, modern features)

if (-not $SkipPwshCheck) {
    $currentVersion = $PSVersionTable.PSVersion
    $requiredMajor = 7
    $requiredMinor = 5

    Write-Host "==================================================================" -ForegroundColor Cyan
    Write-Host "PowerShell Version Check" -ForegroundColor Cyan
    Write-Host "==================================================================" -ForegroundColor Cyan
    Write-Host "Current PowerShell version: $($currentVersion.Major).$($currentVersion.Minor).$($currentVersion.Patch)" -ForegroundColor White
    Write-Host "Required version: $requiredMajor.$requiredMinor or higher" -ForegroundColor White
    Write-Host ""

    if (($currentVersion.Major -lt $requiredMajor) -or
        (($currentVersion.Major -eq $requiredMajor) -and ($currentVersion.Minor -lt $requiredMinor))) {

        Write-Host "[!] PREREQUISITE NOT MET" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "This setup script requires PowerShell 7.5 or higher for optimal functionality:" -ForegroundColor White
        Write-Host "  - Better Unicode support (French characters, emojis)" -ForegroundColor Gray
        Write-Host "  - Improved error handling and debugging" -ForegroundColor Gray
        Write-Host "  - Modern PowerShell features and performance" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Your current version: PowerShell $($currentVersion.Major).$($currentVersion.Minor)" -ForegroundColor Yellow
        Write-Host "Required version: PowerShell 7.5 or higher" -ForegroundColor Green
        Write-Host ""

        # Ask user if they want to install PowerShell 7.5
        if ($y) {
            # Non-interactive mode - auto-install
            $installChoice = "Y"
            Write-Host "[AUTO] Non-interactive mode: Installing PowerShell 7.5 automatically..." -ForegroundColor Cyan
        } else {
            # Interactive mode - ask user
            $installChoice = Read-Host "Would you like to install PowerShell 7.5 now? (Y/N)"
        }

        if ($installChoice -eq "Y" -or $installChoice -eq "y") {
            Write-Host ""
            Write-Host "[+] Installing PowerShell 7.5 using winget..." -ForegroundColor Green

            try {
                # Check if winget is available
                $wingetPath = Get-Command winget -ErrorAction SilentlyContinue

                if (-not $wingetPath) {
                    Write-Host "[ERROR] winget is not available on this system." -ForegroundColor Red
                    Write-Host "Please install PowerShell 7.5 manually from:" -ForegroundColor Yellow
                    Write-Host "  https://github.com/PowerShell/PowerShell/releases" -ForegroundColor Cyan
                    Write-Host ""
                    Write-Host "Or use the Microsoft Store:" -ForegroundColor Yellow
                    Write-Host "  ms-windows-store://pdp/?productid=9MZ1SNWT0N5D" -ForegroundColor Cyan
                    exit 1
                }

                # Install PowerShell 7.5 using winget
                Write-Host "[+] Running: winget install Microsoft.PowerShell --version 7.5.0 --silent" -ForegroundColor Gray
                winget install Microsoft.PowerShell --version 7.5.0 --silent --accept-source-agreements --accept-package-agreements

                if ($LASTEXITCODE -ne 0) {
                    # Try without specific version
                    Write-Host "[+] Trying latest PowerShell 7.x version..." -ForegroundColor Yellow
                    winget install Microsoft.PowerShell --silent --accept-source-agreements --accept-package-agreements
                }

                Write-Host "[OK] PowerShell 7.5 installed successfully!" -ForegroundColor Green
                Write-Host ""
                Write-Host "[+] Launching new PowerShell 7 window and restarting setup..." -ForegroundColor Green
                Write-Host ""

                # Get the script path
                $scriptPath = $MyInvocation.MyCommand.Path

                # Build arguments to pass through
                $args = @("-NoExit", "-File", "`"$scriptPath`"", "-SkipPwshCheck")
                if ($y) {
                    $args += "-y"
                }

                # Launch PowerShell 7 with the script
                Start-Process -FilePath "pwsh" -ArgumentList $args -Wait

                Write-Host "[DONE] Setup completed in new PowerShell 7 window." -ForegroundColor Green
                Write-Host "You can close this window now." -ForegroundColor Gray
                exit 0

            } catch {
                Write-Host "[ERROR] Failed to install PowerShell 7.5" -ForegroundColor Red
                Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host ""
                Write-Host "Please install PowerShell 7.5 manually from:" -ForegroundColor Yellow
                Write-Host "  https://github.com/PowerShell/PowerShell/releases" -ForegroundColor Cyan
                Write-Host ""
                Write-Host "After installation, run this script again." -ForegroundColor Yellow
                exit 1
            }
        } else {
            Write-Host ""
            Write-Host "[!] PowerShell 7.5 installation declined." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "The setup will continue, but you may experience issues:" -ForegroundColor Yellow
            Write-Host "  - Unicode characters may not display correctly" -ForegroundColor Gray
            Write-Host "  - Some features may not work as expected" -ForegroundColor Gray
            Write-Host ""
            Write-Host "To install PowerShell 7.5 later, visit:" -ForegroundColor White
            Write-Host "  https://github.com/PowerShell/PowerShell/releases" -ForegroundColor Cyan
            Write-Host ""

            if (-not $y) {
                $continueChoice = Read-Host "Do you want to continue with the current version? (Y/N)"
                if ($continueChoice -ne "Y" -and $continueChoice -ne "y") {
                    Write-Host "[EXIT] Setup cancelled." -ForegroundColor Red
                    exit 1
                }
            }
        }
    } else {
        Write-Host "[OK] PowerShell version check passed!" -ForegroundColor Green
        Write-Host "==================================================================" -ForegroundColor Cyan
        Write-Host ""
    }
}

# Function to convert string to valid Python module name
function ConvertTo-PythonModuleName {
    param([string]$name)
    # Convert to lowercase, replace spaces and hyphens with underscores, remove invalid chars
    $moduleName = $name.ToLower() -replace '[\s-]', '_' -replace '[^a-z0-9_]', ''
    # Ensure it doesn't start with a number
    if ($moduleName -match '^[0-9]') {
        $moduleName = "module_$moduleName"
    }
    # Capitalize first letter for class-style naming
    return (Get-Culture).TextInfo.ToTitleCase($moduleName)
}

# Function to get current module name from pyproject.toml
function Get-CurrentModuleName {
    if (Test-Path "pyproject.toml") {
        $content = Get-Content "pyproject.toml" -Raw
        if ($content -match 'name\s*=\s*"([^"]+)"') {
            return $matches[1]
        }
    }
    return $null
}

Write-Host "[*] Setting up Python development environment with uv..." -ForegroundColor Green

# Check if uv is already installed
try {
    $uvVersion = uv --version 2>$null
    Write-Host "[OK] uv is already installed ($uvVersion)" -ForegroundColor Green
} catch {
    Write-Host "[+] Installing uv..." -ForegroundColor Yellow

    # Install uv using the official PowerShell installer
    try {
        Invoke-RestMethod -Uri "https://astral.sh/uv/install.ps1" | Invoke-Expression

        # Refresh PATH for current session
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "User") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "Machine")

        # Verify installation
        $uvVersion = uv --version 2>$null
        Write-Host "[OK] uv installed successfully ($uvVersion)" -ForegroundColor Green
    } catch {
        Write-Host "[ERROR] Failed to install uv" -ForegroundColor Red
        Write-Host "Please install uv manually from https://docs.astral.sh/uv/getting-started/installation/" -ForegroundColor Red
        exit 1
    }
}

# Infer module name from directory name
$directoryName = Split-Path -Leaf (Get-Location)
$inferredModuleName = ConvertTo-PythonModuleName $directoryName
$currentModuleName = Get-CurrentModuleName

# Determine the module name to use
$moduleName = $null
if ($currentModuleName -and $currentModuleName -ne "MyModule") {
    # Already configured with a custom name, keep it
    $moduleName = $currentModuleName
    Write-Host "[NOTE] Using existing module name: $moduleName" -ForegroundColor Cyan
} elseif ($y) {
    # Non-interactive mode, use inferred name
    $moduleName = $inferredModuleName
    Write-Host "[NOTE] Using inferred module name: $moduleName (non-interactive mode)" -ForegroundColor Cyan
} else {
    # Interactive mode - ask user
    Write-Host "[NOTE] Module name configuration" -ForegroundColor Yellow
    Write-Host "   Current directory: $directoryName" -ForegroundColor White
    Write-Host "   Suggested module name: $inferredModuleName" -ForegroundColor White
    if ($currentModuleName -and $currentModuleName -ne "MyModule") {
        Write-Host "   Current pyproject.toml name: $currentModuleName" -ForegroundColor White
    }
    Write-Host ""
    $userInput = Read-Host "Enter module name (press Enter for '$inferredModuleName')"
    if ([string]::IsNullOrWhiteSpace($userInput)) {
        $moduleName = $inferredModuleName
    } else {
        $moduleName = ConvertTo-PythonModuleName $userInput
    }
    Write-Host "[OK] Using module name: $moduleName" -ForegroundColor Green
}

# Check if we're in an existing project or need to refresh
if ((Test-Path "pyproject.toml") -and (Test-Path "src")) {
    Write-Host "[REFRESH] Existing project detected. Refreshing..." -ForegroundColor Yellow

    # Remove existing virtual environment if it exists
    if (Test-Path ".venv") {
        Write-Host "[CLEAN] Removing existing virtual environment..." -ForegroundColor Yellow
        Remove-Item -Path ".venv" -Recurse -Force
    }

    # Clean up any cached files
    Write-Host "[CLEAN] Cleaning up cached files..." -ForegroundColor Yellow
    Get-ChildItem -Path . -Recurse -Name "__pycache__" -ErrorAction SilentlyContinue | ForEach-Object {
        Remove-Item -Path $_ -Recurse -Force -ErrorAction SilentlyContinue
    }
    Get-ChildItem -Path . -Recurse -Filter "*.pyc" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
} else {
    Write-Host "[NEW] Setting up new project structure..." -ForegroundColor Yellow
}

# Update module name in pyproject.toml if it's different
if ((Test-Path "pyproject.toml") -and $moduleName -ne $currentModuleName) {
    Write-Host "[CONFIG] Updating module name in pyproject.toml..." -ForegroundColor Yellow
    $content = Get-Content "pyproject.toml" -Raw
    $content = $content -replace 'name\s*=\s*"[^"]+"', "name = `"$moduleName`""
    $content | Set-Content "pyproject.toml" -NoNewline
}

# Sync src directory structure with pyproject.toml name
if (Test-Path "src") {
    # Get the final module name (could be from pyproject.toml if manually edited)
    $finalModuleName = Get-CurrentModuleName
    
    # Find existing module directories in src/ (excluding __pycache__)
    $existingDirs = Get-ChildItem -Path "src" -Directory | Where-Object { $_.Name -ne "__pycache__" }
    
    if ($existingDirs.Count -gt 0) {
        $currentDir = $existingDirs[0].Name

        if ($currentDir -ne $finalModuleName) {
            Write-Host "[CONFIG] Syncing module directory: $currentDir -> $finalModuleName..." -ForegroundColor Yellow

            if (Test-Path "src/$finalModuleName") {
                # Target already exists, remove the old one
                Remove-Item -Path "src/$currentDir" -Recurse -Force
                Write-Host "   Removed duplicate directory: src/$currentDir" -ForegroundColor Gray
            } else {
                # Rename to match pyproject.toml
                Rename-Item -Path "src/$currentDir" -NewName $finalModuleName
                Write-Host "   Renamed: src/$currentDir -> src/$finalModuleName" -ForegroundColor Green
            }
        }
    }

    # Update the moduleName variable for later use
    $moduleName = $finalModuleName
}

# Create virtual environment and install dependencies
Write-Host "[+] Creating virtual environment and installing dependencies..." -ForegroundColor Yellow
try {
    uv sync
} catch {
    Write-Host "[ERROR] Failed to sync dependencies. Make sure you have a valid pyproject.toml file." -ForegroundColor Red
    exit 1
}

# Show virtual environment info
Write-Host "[OK] Virtual environment ready!" -ForegroundColor Green
Write-Host "To activate the environment, run: .venv\Scripts\Activate.ps1 (PowerShell) or .venv\Scripts\activate.bat (Command Prompt)" -ForegroundColor Cyan

# Check for requirements.txt and migrate if found
if (Test-Path "requirements.txt") {
    Write-Host "[+] Found requirements.txt - migrating to uv..." -ForegroundColor Yellow
    try {
        uv add --requirements requirements.txt
        Write-Host "[OK] Successfully migrated requirements.txt to pyproject.toml" -ForegroundColor Green
        Write-Host "[TIP] You can now delete requirements.txt (optional)" -ForegroundColor Cyan
    } catch {
        Write-Host "[!] Could not migrate requirements.txt - you may need to add dependencies manually:" -ForegroundColor Yellow
        Write-Host "    uv add --requirements requirements.txt" -ForegroundColor Cyan
    }
}

# Show installed packages
Write-Host "[+] Installed packages:" -ForegroundColor Yellow
try {
    uv pip list
} catch {
    Write-Host "Could not list packages. Virtual environment may not be properly configured." -ForegroundColor Red
}

Write-Host "[DONE] Setup complete! Your Python development environment is ready." -ForegroundColor Green
Write-Host "[TEST] Run tests with: uv run pytest" -ForegroundColor Cyan
Write-Host "[DEV] Start developing in the src/$moduleName directory" -ForegroundColor Cyan

# Handle git repository initialization
Write-Host ""
if (Test-Path ".git") {
    Write-Host "[GIT] Git repository detected." -ForegroundColor Yellow
    Write-Host "[!] To use this as a new project, you should remove the boilerplate git history." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Options:" -ForegroundColor White
    Write-Host "  1. Remove git history and start fresh (git init)" -ForegroundColor White
    Write-Host "  2. Remove git history, no git init (for git clone later)" -ForegroundColor White
    Write-Host "  3. Keep current git history" -ForegroundColor White
    Write-Host "  4. Skip for now (you can run this script again later)" -ForegroundColor White
    Write-Host ""
    $gitChoice = Read-Host "Choose option (1/2/3/4)"

    switch ($gitChoice) {
        "1" {
            Write-Host "[CLEAN] Removing boilerplate git history..." -ForegroundColor Yellow
            Remove-Item -Path ".git" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "[NEW] Initializing new git repository..." -ForegroundColor Green
            git init
            git add .
            Write-Host "[OK] Ready for your initial commit!" -ForegroundColor Green
            Write-Host "    Run: git commit -m 'Initial commit'" -ForegroundColor Cyan
            Write-Host "    Then add your remote: git remote add origin <your-repo-url>" -ForegroundColor Cyan
        }
        "2" {
            Write-Host "[CLEAN] Removing boilerplate git history..." -ForegroundColor Yellow
            Remove-Item -Path ".git" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "[OK] Ready for git clone! You can now:" -ForegroundColor Green
            Write-Host "    git clone <your-repo-url> temp" -ForegroundColor Cyan
            Write-Host "    Move-Item temp\\.git . && Remove-Item temp -Recurse -Force" -ForegroundColor Cyan
            Write-Host "    git reset --hard HEAD  # to sync with remote" -ForegroundColor Cyan
            Write-Host "    Or initialize later with: git init" -ForegroundColor Cyan
        }
        "3" {
            Write-Host "[OK] Keeping current git history." -ForegroundColor Green
        }
        "4" {
            Write-Host "[SKIP] Skipped git setup. You can run this script again later." -ForegroundColor Yellow
        }
        default {
            Write-Host "[ERROR] Invalid choice. Skipping git setup." -ForegroundColor Red
        }
    }
} else {
    Write-Host "[NOTE] No git repository found. You can:" -ForegroundColor Yellow
    Write-Host "    Initialize: git init && git add . && git commit -m 'Initial commit'" -ForegroundColor Cyan
    Write-Host "    Or clone: git clone <your-repo-url> temp" -ForegroundColor Cyan
    Write-Host "    Then: Move-Item temp\\.git . && Remove-Item temp -Recurse -Force" -ForegroundColor Cyan
    Write-Host "    Finally: git reset --hard HEAD" -ForegroundColor Cyan
}

# Add project setup files and documentation to .gitignore
Write-Host ""
Write-Host "[CONFIG] Adding project setup files to .gitignore..." -ForegroundColor Yellow
$gitignoreEntries = @"

# Project setup files (added by ProjectSetup scripts)
ProjectSetup.ps1
ProjectSetup.sh
HOW2USEIT.md
uv.lock
"@

if (Test-Path ".gitignore") {
    # Check if entries are already there
    $gitignoreContent = Get-Content ".gitignore" -Raw
    if (-not $gitignoreContent.Contains("ProjectSetup.ps1")) {
        Add-Content -Path ".gitignore" -Value $gitignoreEntries
        Write-Host "   [OK] Added setup files to .gitignore" -ForegroundColor Green
    } else {
        Write-Host "   [OK] Setup files already in .gitignore" -ForegroundColor Green
    }
} else {
    Write-Host "   [!] No .gitignore found - setup files not excluded from git" -ForegroundColor Yellow
}

# Unblock all PowerShell scripts to prevent execution policy warnings
Write-Host ""
Write-Host "[CONFIG] Unblocking PowerShell scripts..." -ForegroundColor Yellow
try {
    $ps1Files = Get-ChildItem -Path . -Filter "*.ps1" -File -ErrorAction SilentlyContinue
    $unblockedCount = 0

    foreach ($file in $ps1Files) {
        try {
            Unblock-File -Path $file.FullName -ErrorAction SilentlyContinue
            $unblockedCount++
        } catch {
            # Silently continue if file is already unblocked or can't be unblocked
        }
    }

    if ($unblockedCount -gt 0) {
        Write-Host "   [OK] Unblocked $unblockedCount PowerShell script(s)" -ForegroundColor Green
        Write-Host "   [NOTE] You will no longer be prompted to confirm execution" -ForegroundColor Cyan
    } else {
        Write-Host "   [OK] All PowerShell scripts already unblocked" -ForegroundColor Green
    }
} catch {
    Write-Host "   [!] Could not unblock some scripts - you may be prompted during execution" -ForegroundColor Yellow
}