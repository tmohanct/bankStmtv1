param(
    [string]$OutputDir = 'dist',
    [switch]$IncludeInputPdfs,
    [switch]$IncludeOutputFiles
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSCommandPath
Set-Location $repoRoot

function Copy-FilteredTree {
    param(
        [string]$SourceDir,
        [string]$DestinationDir
    )

    if (-not (Test-Path $SourceDir)) {
        return
    }

    New-Item -ItemType Directory -Path $DestinationDir -Force | Out-Null

    foreach ($entry in Get-ChildItem $SourceDir -Force) {
        if ($entry.PSIsContainer) {
            if ($entry.Name -in @('__pycache__', 'logs', 'output', 'input')) {
                continue
            }

            Copy-FilteredTree -SourceDir $entry.FullName -DestinationDir (Join-Path $DestinationDir $entry.Name)
            continue
        }

        if ($entry.Extension -in @('.pyc', '.pyo')) {
            continue
        }

        Copy-Item $entry.FullName -Destination (Join-Path $DestinationDir $entry.Name) -Force
    }
}

$distRoot = Join-Path $repoRoot $OutputDir
New-Item -ItemType Directory -Path $distRoot -Force | Out-Null

$timestamp = Get-Date -Format 'yyMMdd_HHmmss'
$packageName = "bankStmtv1_windows_$timestamp"
$stageDir = Join-Path $distRoot $packageName
$zipPath = Join-Path $distRoot "$packageName.zip"

if (Test-Path $stageDir) {
    Remove-Item $stageDir -Recurse -Force
}

if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

New-Item -ItemType Directory -Path $stageDir -Force | Out-Null

$rootFiles = @(
    'AGENTS.md',
    'README.md',
    'SETUP_WINDOWS.md',
    'requirements.txt',
    'run.py',
    'run_bank_parser.bat',
    'setup_windows.bat',
    'setup_windows.ps1',
    'install_fresh_machine.bat'
)

foreach ($file in $rootFiles) {
    $sourcePath = Join-Path $repoRoot $file
    if (Test-Path $sourcePath) {
        Copy-Item $sourcePath -Destination (Join-Path $stageDir $file) -Force
    }
}

Copy-FilteredTree -SourceDir (Join-Path $repoRoot 'src') -DestinationDir (Join-Path $stageDir 'src')

$inputDir = Join-Path $stageDir 'input'
$outputDir = Join-Path $stageDir 'output'
New-Item -ItemType Directory -Path $inputDir -Force | Out-Null
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$rulesCandidates = @(
    (Join-Path $repoRoot 'input\Rules.xlsx'),
    (Join-Path $repoRoot 'src\input\Rules.xlsx')
)
$rulesSource = $rulesCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($rulesSource) {
    Copy-Item $rulesSource -Destination (Join-Path $inputDir 'Rules.xlsx') -Force
}
else {
    Write-Warning 'Rules.xlsx was not found. Add input\Rules.xlsx before sharing this package.'
}

if ($IncludeInputPdfs) {
    $repoInputDir = Join-Path $repoRoot 'input'
    if (Test-Path $repoInputDir) {
        Get-ChildItem $repoInputDir -File -Filter '*.pdf' | ForEach-Object {
            Copy-Item $_.FullName -Destination (Join-Path $inputDir $_.Name) -Force
        }
    }
}

if ($IncludeOutputFiles) {
    $repoOutputDir = Join-Path $repoRoot 'output'
    if (Test-Path $repoOutputDir) {
        Get-ChildItem $repoOutputDir -File | ForEach-Object {
            Copy-Item $_.FullName -Destination (Join-Path $outputDir $_.Name) -Force
        }
    }
}

Compress-Archive -Path $stageDir -DestinationPath $zipPath -Force

Write-Host "Windows package created: $zipPath"
Write-Host 'Default package contents: source code, setup scripts, and input\Rules.xlsx.'
Write-Host 'Not included by default: .venv, logs, output files, and input PDFs.'
