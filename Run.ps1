Start-Transcript .\lastlogexecution.log
try {

    .\Scripts\Activate.ps1

    $Confirm = $( Write-Host "Run of the Current Pipe refresh [Y]/N :" -NoNewline -ForegroundColor Green; Read-Host)
    $Confirm = $Confirm.ToLower()
    if ($confirm -ne 'n') {
        Write-Host "Running ..." -ForegroundColor Yellow 
        python UpdatePipe.py
    } else {
    Write-Host "Skipping ..." -ForegroundColor Yellow 
    }

    $Confirm = $( Write-Host "Run of the End User Pipe refresh [Y]/N :" -NoNewline -ForegroundColor Green; Read-Host)
    $Confirm = $Confirm.ToLower()
    if ($confirm -ne 'n') {
        Write-Host "Running ..." -ForegroundColor Yellow 
        python UpdateEndUser.py
    } else {
    Write-Host "Skipping ..." -ForegroundColor Yellow 
    }

}
catch {
    $Error[0].Exception
} 
finally {
    Stop-Transcript
}