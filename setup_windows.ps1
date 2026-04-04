param(
    [switch]$SkipTesseract,
    [switch]$ForceRecreateVenv
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSCommandPath
Set-Location $repoRoot

function Invoke-Checked {
    param([string[]]$Command)

    $display = $Command -join ' '
    Write-Host $display

    $exe = $Command[0]
    if ($Command.Count -gt 1) {
        $args = $Command[1..($Command.Count - 1)]
    }
    else {
        $args = @()
    }

    & $exe @args
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed: $display"
    }
}

function Read-PyVenvConfig {
    param([string]$ConfigPath)

    $config = @{}
    if (-not (Test-Path $ConfigPath)) {
        return $config
    }

    foreach ($line in Get-Content $ConfigPath) {
        if ($line -match '^\s*([^=]+?)\s*=\s*(.+?)\s*$') {
            $config[$matches[1].Trim().ToLowerInvariant()] = $matches[2].Trim()
        }
    }

    return $config
}

function Get-PythonLauncher {
    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCmd) {
        return @($pythonCmd.Source)
    }

    $pyCmd = Get-Command py -ErrorAction SilentlyContinue
    if ($pyCmd) {
        return @($pyCmd.Source, '-3')
    }

    throw 'Python 3 was not found. Install Python 3.11 or newer and add it to PATH.'
}

function Get-VenvHealth {
    param([string]$VenvDir)

    $venvPython = Join-Path $VenvDir 'Scripts\python.exe'
    if (-not (Test-Path $venvPython)) {
        return [pscustomobject]@{
            Exists = $false
            Healthy = $false
            Reason = 'Virtual environment was not found.'
        }
    }

    $configPath = Join-Path $VenvDir 'pyvenv.cfg'
    if (-not (Test-Path $configPath)) {
        return [pscustomobject]@{
            Exists = $true
            Healthy = $false
            Reason = 'pyvenv.cfg is missing.'
        }
    }

    $config = Read-PyVenvConfig -ConfigPath $configPath
    $baseExecutable = $config['executable']
    if ($baseExecutable -and -not (Test-Path $baseExecutable)) {
        return [pscustomobject]@{
            Exists = $true
            Healthy = $false
            Reason = "Base interpreter is missing: $baseExecutable"
        }
    }

    $baseHome = $config['home']
    if ($baseHome -and -not (Test-Path $baseHome)) {
        return [pscustomobject]@{
            Exists = $true
            Healthy = $false
            Reason = "Base Python home is missing: $baseHome"
        }
    }

    try {
        & $venvPython --version | Out-Null
        if ($LASTEXITCODE -ne 0) {
            return [pscustomobject]@{
                Exists = $true
                Healthy = $false
                Reason = "python.exe returned exit code $LASTEXITCODE."
            }
        }
    }
    catch {
        return [pscustomobject]@{
            Exists = $true
            Healthy = $false
            Reason = $_.Exception.Message
        }
    }

    return [pscustomobject]@{
        Exists = $true
        Healthy = $true
        Reason = ''
    }
}

function Get-TesseractPath {
    $tesseractCmd = Get-Command tesseract -ErrorAction SilentlyContinue
    if ($tesseractCmd) {
        return $tesseractCmd.Source
    }

    $candidates = @(
        'C:\Program Files\Tesseract-OCR\tesseract.exe',
        'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    return $null
}

Write-Host '==> Checking Python'
$pythonLauncher = @(Get-PythonLauncher)
Invoke-Checked (@($pythonLauncher) + @('--version'))

$venvDir = Join-Path $repoRoot '.venv'
$venvPython = Join-Path $venvDir 'Scripts\python.exe'
$venvHealth = Get-VenvHealth -VenvDir $venvDir

if ($ForceRecreateVenv -and (Test-Path $venvDir)) {
    Write-Host ''
    Write-Host '==> Removing existing virtual environment because -ForceRecreateVenv was requested'
    Remove-Item $venvDir -Recurse -Force
    $venvHealth = Get-VenvHealth -VenvDir $venvDir
}
elseif ($venvHealth.Exists -and -not $venvHealth.Healthy) {
    Write-Host ''
    Write-Warning "Existing .venv is stale or broken. Recreating it. Reason: $($venvHealth.Reason)"
    Remove-Item $venvDir -Recurse -Force
    $venvHealth = Get-VenvHealth -VenvDir $venvDir
}

if (-not (Test-Path $venvPython)) {
    Write-Host ''
    Write-Host '==> Creating virtual environment'
    Invoke-Checked (@($pythonLauncher) + @('-m', 'venv', '.venv'))
}
else {
    Write-Host ''
    Write-Host '==> Reusing healthy virtual environment'
}

Write-Host ''
Write-Host '==> Installing Python packages'
Invoke-Checked @($venvPython, '-m', 'pip', 'install', '--upgrade', 'pip')
Invoke-Checked @($venvPython, '-m', 'pip', 'install', '-r', 'requirements.txt')

if (-not $SkipTesseract) {
    Write-Host ''
    Write-Host '==> Checking Tesseract OCR'
    $tesseractPath = Get-TesseractPath

    if (-not $tesseractPath) {
        $wingetCmd = Get-Command winget -ErrorAction SilentlyContinue
        if ($wingetCmd) {
            Write-Host 'Tesseract not found. Installing with winget because ICICI parser needs OCR.'
            Invoke-Checked @(
                $wingetCmd.Source,
                'install',
                '--id',
                'UB-Mannheim.TesseractOCR',
                '-e',
                '--accept-package-agreements',
                '--accept-source-agreements'
            )
            $tesseractPath = Get-TesseractPath
        }
        else {
            Write-Warning 'winget was not found. Install Tesseract manually if you need ICICI statement support.'
        }
    }

    if ($tesseractPath) {
        [Environment]::SetEnvironmentVariable('TESSERACT_CMD', $tesseractPath, 'User')
        Write-Host "Tesseract found: $tesseractPath"
        Write-Host 'Saved TESSERACT_CMD for the current Windows user.'
    }
    else {
        Write-Warning 'Tesseract is still not available. Axis, CUB, HDFC, IDBI, and SBI can run, but ICICI will not work until Tesseract is installed.'
    }
}
else {
    Write-Host ''
    Write-Host '==> Skipping Tesseract setup'
}

Write-Host ''
Write-Host '==> Verifying parser files'
$pythonFiles = @(
    (Join-Path $repoRoot 'run.py')
)
$pythonFiles += Get-ChildItem (Join-Path $repoRoot 'src') -Recurse -Filter '*.py' | ForEach-Object {
    $_.FullName
}
Invoke-Checked (@($venvPython, '-m', 'py_compile') + $pythonFiles)

Write-Host ''
Write-Host 'Setup complete.'
Write-Host 'Use these commands from the repo root:'
Write-Host '.\install_fresh_machine.bat'
Write-Host '.\run_bank_parser.bat --bank sbi --file sbi.pdf'
Write-Host '.\run_bank_parser.bat --bank axis --file "axis.pdf;axis2.pdf"'
Write-Host ''
Write-Host 'Output files are written to .\output'
