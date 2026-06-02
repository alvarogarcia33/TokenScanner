param(
    [string]$IsccPath = "",
    [string]$InstallerScript = ".\installer\TokenScannerLocal.iss"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $InstallerScript)) {
    throw "No se encontró el script del instalador: $InstallerScript"
}

$candidatePaths = @()
if ($IsccPath) {
    $candidatePaths += $IsccPath
}
$candidatePaths += @(
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe",
    (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe")
)

$resolvedIscc = $null
foreach ($candidate in $candidatePaths) {
    if (Test-Path $candidate) {
        $resolvedIscc = $candidate
        break
    }
}

if (-not $resolvedIscc) {
    $command = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($command) {
        $resolvedIscc = $command.Source
    }
}

if (-not $resolvedIscc) {
    throw "No se encontró ISCC.exe. Instalá Inno Setup o pasá -IsccPath."
}

& $resolvedIscc $InstallerScript
