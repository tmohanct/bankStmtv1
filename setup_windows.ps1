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

function Get-InstalledPythonExeCandidates {
    $candidates = New-Object System.Collections.Generic.List[string]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    $registryRoots = @(
        'HKCU:\Software\Python\PythonCore',
        'HKLM:\Software\Python\PythonCore',
        'HKLM:\Software\WOW6432Node\Python\PythonCore'
    )

    foreach ($root in $registryRoots) {
        if (-not (Test-Path $root)) {
            continue
        }

        foreach ($versionKey in Get-ChildItem $root -ErrorAction SilentlyContinue) {
            $installPathKey = Join-Path $versionKey.PSPath 'InstallPath'
            if (-not (Test-Path $installPathKey)) {
                continue
            }

            try {
                $installKey = Get-Item $installPathKey -ErrorAction Stop
            }
            catch {
                continue
            }

            $defaultDir = $installKey.GetValue('')
            if ($defaultDir) {
                $pythonExe = Join-Path $defaultDir 'python.exe'
                if ((Test-Path $pythonExe) -and $seen.Add($pythonExe)) {
                    [void]$candidates.Add($pythonExe)
                }
            }

            $executablePath = $installKey.GetValue('ExecutablePath')
            if ($executablePath -and (Test-Path $executablePath) -and $seen.Add($executablePath)) {
                [void]$candidates.Add($executablePath)
            }
        }
    }

    return $candidates.ToArray()
}

function Test-PythonCommand {
    param([string[]]$Command)

    $exePath = $Command[0]

    if ($exePath -match '\\WindowsApps\\python(?:3(?:\.\d+)?)?\.exe$') {
        return [pscustomobject]@{
            Usable = $false
            IsStoreAlias = $true
            Reason = 'Found the Microsoft Store alias instead of a real Python install.'
            VersionText = ''
        }
    }

    if ($Command.Count -gt 1) {
        $args = @($Command[1..($Command.Count - 1)] + @('--version'))
    }
    else {
        $args = @('--version')
    }

    try {
        $output = & $exePath @args 2>&1
        $exitCode = $LASTEXITCODE
    }
    catch {
        return [pscustomobject]@{
            Usable = $false
            IsStoreAlias = $false
            Reason = $_.Exception.Message
            VersionText = ''
        }
    }

    if ($exitCode -ne 0) {
        $firstLine = @($output)[0]
        $reason = if ($firstLine) { [string]$firstLine } else { "Exit code $exitCode" }
        return [pscustomobject]@{
            Usable = $false
            IsStoreAlias = $false
            Reason = $reason.Trim()
            VersionText = ''
        }
    }

    $versionText = [string](@($output)[0])
    $versionText = $versionText.Trim()
    if ($versionText -notmatch '^Python\s+(\d+)\.(\d+)') {
        return [pscustomobject]@{
            Usable = $false
            IsStoreAlias = $false
            Reason = "Unexpected version output: $versionText"
            VersionText = $versionText
        }
    }

    $major = [int]$matches[1]
    $minor = [int]$matches[2]
    if ($major -lt 3 -or ($major -eq 3 -and $minor -lt 11)) {
        return [pscustomobject]@{
            Usable = $false
            IsStoreAlias = $false
            Reason = "Python 3.11 or newer is required. Found $versionText."
            VersionText = $versionText
        }
    }

    return [pscustomobject]@{
        Usable = $true
        IsStoreAlias = $false
        Reason = ''
        VersionText = $versionText
    }
}

function Get-PythonLauncher {
    param([switch]$AllowMissing)

    $candidateCommands = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $sawStoreAlias = $false

    function Add-PythonCandidate {
        param([string[]]$Command)

        $key = $Command -join "`0"
        if ($seen.Add($key)) {
            [void]$candidateCommands.Add($Command)
        }
    }

    $pyCmd = Get-Command py -ErrorAction SilentlyContinue
    if ($pyCmd) {
        Add-PythonCandidate @($pyCmd.Source, '-3')
    }

    foreach ($commandName in @('python', 'python3')) {
        foreach ($command in @(Get-Command $commandName -All -ErrorAction SilentlyContinue)) {
            if ($command.Source) {
                Add-PythonCandidate @($command.Source)
            }
        }
    }

    foreach ($candidatePath in Get-InstalledPythonExeCandidates) {
        Add-PythonCandidate @($candidatePath)
    }

    foreach ($candidate in $candidateCommands) {
        $result = Test-PythonCommand -Command $candidate
        if ($result.Usable) {
            return $candidate
        }

        if ($result.IsStoreAlias) {
            $sawStoreAlias = $true
        }
    }

    if ($AllowMissing) {
        return $null
    }

    if ($sawStoreAlias) {
        throw 'Python 3.11 or newer is not installed. Windows only found the Microsoft Store python alias.'
    }

    throw 'Python 3.11 or newer was not found. Install it and add it to PATH, or install the Windows py launcher.'
}

function Ensure-PythonLauncher {
    $pythonLauncher = Get-PythonLauncher -AllowMissing
    if ($pythonLauncher) {
        return $pythonLauncher
    }

    $wingetCmd = Get-Command winget -ErrorAction SilentlyContinue
    if ($wingetCmd) {
        Write-Host 'Python 3.11+ was not found. Installing Python 3.11 with winget.'
        Invoke-Checked @(
            $wingetCmd.Source,
            'install',
            '--id',
            'Python.Python.3.11',
            '-e',
            '--accept-package-agreements',
            '--accept-source-agreements'
        )

        $pythonLauncher = Get-PythonLauncher -AllowMissing
        if ($pythonLauncher) {
            return $pythonLauncher
        }
    }

    throw 'Python 3.11 or newer is required. Install it from python.org for Windows and rerun this setup script.'
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
$pythonLauncher = @(Ensure-PythonLauncher)
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
Write-Host '.\install_new_machine.bat'
Write-Host '.\run_bank_parser.bat --bank sbi --file sbi.pdf'
Write-Host '.\run_bank_parser.bat --bank axis --file "axis.pdf;axis2.pdf"'
Write-Host ''
Write-Host 'Output files are written to .\output'
