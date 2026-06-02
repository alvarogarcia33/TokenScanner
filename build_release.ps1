param(
    [string]$PythonPath = "",
    [string]$ReleaseDir = "release",
    [string]$ExeName = "TokenScannerLocal"
)

$ErrorActionPreference = "Stop"

$releasePath = Join-Path $PSScriptRoot $ReleaseDir
$distPath = Join-Path $PSScriptRoot "dist_release"
$workPath = Join-Path $PSScriptRoot "build_release_work"

if (Test-Path $releasePath) {
    Get-ChildItem $releasePath -File | Remove-Item -Force
}
else {
    New-Item -ItemType Directory -Path $releasePath | Out-Null
}

if (Test-Path $distPath) {
    Remove-Item $distPath -Recurse -Force
}

if (Test-Path $workPath) {
    Remove-Item $workPath -Recurse -Force
}

& (Join-Path $PSScriptRoot "build_exe.ps1") -PythonPath $PythonPath -ExeName $ExeName -DistPath $distPath -WorkPath $workPath

$builtExe = Join-Path $distPath "$ExeName.exe"
$releaseExe = Join-Path $releasePath "$ExeName.exe"

Copy-Item $builtExe $releaseExe -Force

Write-Host "Release lista en: $releaseExe"
