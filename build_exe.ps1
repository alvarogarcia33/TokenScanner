param(
    [string]$PythonPath = "",
    [string]$ExeName = "TokenScannerLocal",
    [string]$DistPath = "dist",
    [string]$WorkPath = "build"
)

$ErrorActionPreference = "Stop"

if (-not $PythonPath) {
    if (Test-Path ".\.venv313\Scripts\python.exe") {
        $PythonPath = ".\.venv313\Scripts\python.exe"
    }
    elseif (Test-Path ".\.venv\Scripts\python.exe") {
        $PythonPath = ".\.venv\Scripts\python.exe"
    }
    else {
        $PythonPath = "python"
    }
}
elseif (-not (Test-Path $PythonPath)) {
    $PythonPath = "python"
}

$minorVersion = & $PythonPath -c "import sys; print(sys.version_info.minor if sys.version_info.major == 3 else -1)"
if ([int]$minorVersion -ge 14) {
    Write-Warning "PyInstaller pudo generar el .exe con Python 3.14+, pero web3/pydantic emiten warnings. Para releases mas solidos, preferi Python 3.12."
}

& $PythonPath -m pip install -r requirements-build.txt

& $PythonPath -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --name $ExeName `
    --distpath $DistPath `
    --workpath $WorkPath `
    --add-data "templates;templates" `
    --add-data "tokens.json;." `
    --collect-all openpyxl `
    app.py
