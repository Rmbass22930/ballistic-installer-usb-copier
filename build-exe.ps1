[CmdletBinding()]
param(
    [string]$OutputRoot = "$PSScriptRoot\dist"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = $PSScriptRoot
$buildRoot = Join-Path $projectRoot 'build'
$specRoot = Join-Path $projectRoot 'spec'
$outputRoot = [System.IO.Path]::GetFullPath($OutputRoot)
$pythonExe = 'py'
$pythonArgs = @('-3.13')
$stageRoot = Join-Path $projectRoot '_pyi_stage'
$stageDist = Join-Path $stageRoot 'dist'
$stageWork = Join-Path $stageRoot 'build'
$finalExe = Join-Path $outputRoot 'BallisticInstallerUsbCopier.exe'

New-Item -ItemType Directory -Force -Path $outputRoot | Out-Null
New-Item -ItemType Directory -Force -Path $buildRoot | Out-Null
New-Item -ItemType Directory -Force -Path $specRoot | Out-Null
Remove-Item -LiteralPath $stageDist -Force -Recurse -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $stageWork -Force -Recurse -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $stageDist | Out-Null
New-Item -ItemType Directory -Force -Path $stageWork | Out-Null

& $pythonExe @pythonArgs -m PyInstaller `
    --noconfirm `
    --onefile `
    --windowed `
    --name BallisticInstallerUsbCopier `
    --distpath $stageDist `
    --workpath $stageWork `
    --specpath $specRoot `
    (Join-Path $projectRoot 'main.py')

if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller build failed with exit code $LASTEXITCODE"
}

$stagedExe = Join-Path $stageDist 'BallisticInstallerUsbCopier.exe'
if (-not (Test-Path $stagedExe)) {
    throw "Expected staged executable was not created: $stagedExe"
}

Get-Process -Name 'BallisticInstallerUsbCopier' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Copy-Item -Path $stagedExe -Destination $finalExe -Force

Write-Host ''
Write-Host "Build complete: $finalExe" -ForegroundColor Green
