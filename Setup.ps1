Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
choco install python3 -Y
python -m venv .
.\Scripts\Activate.ps1
pip install -r requirement.txt
python.exe -m pip install --upgrade pip