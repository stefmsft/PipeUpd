Start-Transcript .\lastlogexecution.log
try {

.\Scripts\Activate.ps1
python UpdatePipe.py

}
catch {
    $Error[0].Exception
} 
finally {
    Stop-Transcript
}