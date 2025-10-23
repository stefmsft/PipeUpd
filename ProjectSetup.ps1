#!/usr/bin/env pwsh

# Python Boilerplate Setup Script
# This script installs uv and sets up a Python development environment

# Parse command line arguments
param(
    [switch]$y,
    [switch]$help
)

# Set strict mode to exit on any error
$ErrorActionPreference = "Stop"

# Show help if requested
if ($help) {
    Write-Host "Python Boilerplate Setup Script" -ForegroundColor Green
    Write-Host "Usage: .\setup.ps1 [-y] [-help]" -ForegroundColor White
    Write-Host "  -y      : Non-interactive mode (use defaults)" -ForegroundColor White
    Write-Host "  -help   : Show this help message" -ForegroundColor White
    exit 0
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

Write-Host "🚀 Setting up Python development environment with uv..." -ForegroundColor Green

# Check if uv is already installed
try {
    $uvVersion = uv --version 2>$null
    Write-Host "✅ uv is already installed ($uvVersion)" -ForegroundColor Green
} catch {
    Write-Host "📦 Installing uv..." -ForegroundColor Yellow
    
    # Install uv using the official PowerShell installer
    try {
        Invoke-RestMethod -Uri "https://astral.sh/uv/install.ps1" | Invoke-Expression
        
        # Refresh PATH for current session
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "User") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "Machine")
        
        # Verify installation
        $uvVersion = uv --version 2>$null
        Write-Host "✅ uv installed successfully ($uvVersion)" -ForegroundColor Green
    } catch {
        Write-Host "❌ Failed to install uv" -ForegroundColor Red
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
    Write-Host "📝 Using existing module name: $moduleName" -ForegroundColor Cyan
} elseif ($y) {
    # Non-interactive mode, use inferred name
    $moduleName = $inferredModuleName
    Write-Host "📝 Using inferred module name: $moduleName (non-interactive mode)" -ForegroundColor Cyan
} else {
    # Interactive mode - ask user
    Write-Host "📝 Module name configuration" -ForegroundColor Yellow
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
    Write-Host "✅ Using module name: $moduleName" -ForegroundColor Green
}

# Check if we're in an existing project or need to refresh
if ((Test-Path "pyproject.toml") -and (Test-Path "src")) {
    Write-Host "🔄 Existing project detected. Refreshing..." -ForegroundColor Yellow
    
    # Remove existing virtual environment if it exists
    if (Test-Path ".venv") {
        Write-Host "🗑️  Removing existing virtual environment..." -ForegroundColor Yellow
        Remove-Item -Path ".venv" -Recurse -Force
    }
    
    # Clean up any cached files
    Write-Host "🧹 Cleaning up cached files..." -ForegroundColor Yellow
    Get-ChildItem -Path . -Recurse -Name "__pycache__" -ErrorAction SilentlyContinue | ForEach-Object {
        Remove-Item -Path $_ -Recurse -Force -ErrorAction SilentlyContinue
    }
    Get-ChildItem -Path . -Recurse -Filter "*.pyc" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
} else {
    Write-Host "🆕 Setting up new project structure..." -ForegroundColor Yellow
}

# Update module name in pyproject.toml if it's different
if ((Test-Path "pyproject.toml") -and $moduleName -ne $currentModuleName) {
    Write-Host "🔧 Updating module name in pyproject.toml..." -ForegroundColor Yellow
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
            Write-Host "🔧 Syncing module directory: $currentDir → $finalModuleName..." -ForegroundColor Yellow
            
            if (Test-Path "src/$finalModuleName") {
                # Target already exists, remove the old one
                Remove-Item -Path "src/$currentDir" -Recurse -Force
                Write-Host "   Removed duplicate directory: src/$currentDir" -ForegroundColor Gray
            } else {
                # Rename to match pyproject.toml
                Rename-Item -Path "src/$currentDir" -NewName $finalModuleName
                Write-Host "   Renamed: src/$currentDir → src/$finalModuleName" -ForegroundColor Green
            }
        }
    }
    
    # Update the moduleName variable for later use
    $moduleName = $finalModuleName
}

# Create virtual environment and install dependencies
Write-Host "🔧 Creating virtual environment and installing dependencies..." -ForegroundColor Yellow
try {
    uv sync
} catch {
    Write-Host "❌ Failed to sync dependencies. Make sure you have a valid pyproject.toml file." -ForegroundColor Red
    exit 1
}

# Show virtual environment info
Write-Host "🐍 Virtual environment ready!" -ForegroundColor Green
Write-Host "To activate the environment, run: .venv\Scripts\Activate.ps1 (PowerShell) or .venv\Scripts\activate.bat (Command Prompt)" -ForegroundColor Cyan

# Check for requirements.txt and migrate if found
if (Test-Path "requirements.txt") {
    Write-Host "📋 Found requirements.txt - migrating to uv..." -ForegroundColor Yellow
    try {
        uv add --requirements requirements.txt
        Write-Host "✅ Successfully migrated requirements.txt to pyproject.toml" -ForegroundColor Green
        Write-Host "💡 You can now delete requirements.txt (optional)" -ForegroundColor Cyan
    } catch {
        Write-Host "⚠️  Could not migrate requirements.txt - you may need to add dependencies manually:" -ForegroundColor Yellow
        Write-Host "    uv add --requirements requirements.txt" -ForegroundColor Cyan
    }
}

# Show installed packages
Write-Host "📦 Installed packages:" -ForegroundColor Yellow
try {
    uv pip list
} catch {
    Write-Host "Could not list packages. Virtual environment may not be properly configured." -ForegroundColor Red
}

Write-Host "✨ Setup complete! Your Python development environment is ready." -ForegroundColor Green
Write-Host "🧪 Run tests with: uv run pytest" -ForegroundColor Cyan
Write-Host "🚀 Start developing in the src/$moduleName directory" -ForegroundColor Cyan

# Handle git repository initialization
Write-Host ""
if (Test-Path ".git") {
    Write-Host "🔍 Git repository detected." -ForegroundColor Yellow
    Write-Host "⚠️  To use this as a new project, you should remove the boilerplate git history." -ForegroundColor Yellow
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
            Write-Host "🗑️  Removing boilerplate git history..." -ForegroundColor Yellow
            Remove-Item -Path ".git" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "🆕 Initializing new git repository..." -ForegroundColor Green
            git init
            git add .
            Write-Host "📝 Ready for your initial commit!" -ForegroundColor Green
            Write-Host "    Run: git commit -m 'Initial commit'" -ForegroundColor Cyan
            Write-Host "    Then add your remote: git remote add origin <your-repo-url>" -ForegroundColor Cyan
        }
        "2" {
            Write-Host "🗑️  Removing boilerplate git history..." -ForegroundColor Yellow
            Remove-Item -Path ".git" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "📝 Ready for git clone! You can now:" -ForegroundColor Green
            Write-Host "    git clone <your-repo-url> temp" -ForegroundColor Cyan
            Write-Host "    Move-Item temp\\.git . && Remove-Item temp -Recurse -Force" -ForegroundColor Cyan
            Write-Host "    git reset --hard HEAD  # to sync with remote" -ForegroundColor Cyan
            Write-Host "    Or initialize later with: git init" -ForegroundColor Cyan
        }
        "3" {
            Write-Host "✅ Keeping current git history." -ForegroundColor Green
        }
        "4" {
            Write-Host "⏭️  Skipped git setup. You can run this script again later." -ForegroundColor Yellow
        }
        default {
            Write-Host "❓ Invalid choice. Skipping git setup." -ForegroundColor Red
        }
    }
} else {
    Write-Host "📝 No git repository found. You can:" -ForegroundColor Yellow
    Write-Host "    Initialize: git init && git add . && git commit -m 'Initial commit'" -ForegroundColor Cyan
    Write-Host "    Or clone: git clone <your-repo-url> temp" -ForegroundColor Cyan
    Write-Host "    Then: Move-Item temp\\.git . && Remove-Item temp -Recurse -Force" -ForegroundColor Cyan
    Write-Host "    Finally: git reset --hard HEAD" -ForegroundColor Cyan
}

# Add project setup files and documentation to .gitignore
Write-Host ""
Write-Host "🚫 Adding project setup files to .gitignore..." -ForegroundColor Yellow
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
        Write-Host "   ✅ Added setup files to .gitignore" -ForegroundColor Green
    } else {
        Write-Host "   ✅ Setup files already in .gitignore" -ForegroundColor Green
    }
} else {
    Write-Host "   ⚠️  No .gitignore found - setup files not excluded from git" -ForegroundColor Yellow
}