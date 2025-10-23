Start-Transcript .\lastlogexecution-UpdateEndUser.log
try {

    # Check if uv is available
    $uvAvailable = $false
    try {
        $uvVersion = uv --version 2>$null
        if ($LASTEXITCODE -eq 0) {
            $uvAvailable = $true
            Write-Host "Using uv (version: $uvVersion)" -ForegroundColor Green
        }
    }
    catch {
        # uv not found, will fall back to traditional venv
    }

    # Activate environment
    if ($uvAvailable) {
        # Use uv to run the script directly
        Write-Host "Running UpdateEndUser.py with uv..." -ForegroundColor Yellow
        uv run UpdateEndUser.py
    }
    else {
        # Fall back to traditional virtual environment activation
        Write-Host "Using traditional virtual environment..." -ForegroundColor Yellow
        if (Test-Path ".\Scripts\Activate.ps1") {
            .\Scripts\Activate.ps1
        }
        elseif (Test-Path ".\.venv\Scripts\Activate.ps1") {
            .\.venv\Scripts\Activate.ps1
        }
        else {
            Write-Host "Warning: No virtual environment activation script found. Running with system Python." -ForegroundColor Yellow
        }

        Write-Host "Running UpdateEndUser.py..." -ForegroundColor Yellow
        python UpdateEndUser.py
    }

}
catch {
    Write-Host "Error occurred:" -ForegroundColor Red
    Write-Host $Error[0].Exception -ForegroundColor Red
}
finally {
    Stop-Transcript
}