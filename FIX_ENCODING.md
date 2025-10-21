# Fix PowerShell Script Encoding Issues

## Problem
After cloning this repository, PowerShell scripts fail with encoding errors like:
- "Jeton inattendu" (unexpected token)
- Garbled characters in error messages (âœ…, etc.)

## Root Cause
PowerShell scripts contain Unicode emoji characters that require UTF-8 encoding with BOM (Byte Order Mark). Git may not apply the correct encoding on initial clone.

## Solution

Run these commands in PowerShell from the project directory:

```powershell
# Step 1: Remove all files from git cache
git rm --cached -r .

# Step 2: Re-add all files with correct encoding
git reset --hard HEAD

# Step 3: Verify the fix worked
.\ProjectSetup.ps1 -help
```

If you still see errors, manually convert the file encoding:

```powershell
# Convert ProjectSetup.ps1 to UTF-8 with BOM
$content = Get-Content .\ProjectSetup.ps1 -Raw
$utf8BOM = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText("$PWD\ProjectSetup.ps1", $content, $utf8BOM)

# Test again
.\ProjectSetup.ps1 -help
```

## Prevention

The `.gitattributes` file has been added to ensure correct encoding for all PowerShell scripts. This fix should only be needed once per cloned repository.

## Files Affected
- `ProjectSetup.ps1`
- `UpdatePipe.ps1`
- `UpdateEndUser.ps1`
- `Setup.ps1`
- `Run.ps1`

All `.ps1` files use UTF-8 with BOM encoding for Unicode emoji support.
