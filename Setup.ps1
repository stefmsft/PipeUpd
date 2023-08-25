Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
choco install git -Y
choco install python3 -Y
try {
    python --version
}
catch {
    RefreshEnv
    write-output ("")
    write-output ("####### Relancez la commande Setup.ps1 s'il vous plait #######")
    exit
}
if (!(Test-Path "pyvenv.cfg")) {python -m venv .}
.\Scripts\Activate.ps1
pip install -r requirement.txt
python.exe -m pip install --upgrade pip
if (!(Test-Path ".env")) { copy-item .env.template -Destination .env}

write-output ("")
write-output ("Next Steps :")
write-output ("     - Configuration du fichier .env")
write-output ("     - Execution du script de mise a jour du Pipe (Run.ps1)")