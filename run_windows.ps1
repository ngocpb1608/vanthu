$ErrorActionPreference = "Stop"
$HERE = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $HERE
$py = "python"; try { $null = Get-Command python -ErrorAction Stop } catch { $py = "python3" }
& $py -m venv .venv
& ".\.venv\Scripts\Activate.ps1"
& $py -m pip install --upgrade pip
pip install -r requirements.txt
if (-not $env:SECRET_KEY) { $env:SECRET_KEY = "something-strong" }
if (-not $env:DATABASE_URL) { $env:DATABASE_URL = "sqlite:///data.db" }
& $py -m flask --app server run --debug --port 5050
