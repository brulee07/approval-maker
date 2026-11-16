param([switch]$Install)
$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot

function Find-Python {
    $localPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
    if (Test-Path -LiteralPath $localPython) { return $localPython }
    $runtimeRecord = Join-Path $PSScriptRoot '.runtime\python-path.txt'
    if (Test-Path -LiteralPath $runtimeRecord) {
        $recordedPython = (Get-Content -LiteralPath $runtimeRecord -Raw).Trim()
        if (Test-Path -LiteralPath $recordedPython) { return $recordedPython }
    }
    $packages = Get-AppxPackage -Name 'OpenAI.CodexPrimaryRuntime*' -ErrorAction SilentlyContinue | Sort-Object Version -Descending
    foreach ($package in $packages) {
        $packagePython = Join-Path $package.InstallLocation 'dependencies\python\python.exe'
        if (Test-Path -LiteralPath $packagePython) { return $packagePython }
    }
    $command = Get-Command python -ErrorAction SilentlyContinue
    if ($command -and $command.Source -notlike '*\Microsoft\WindowsApps\*') { return $command.Source }
    $launcher = Get-Command py -ErrorAction SilentlyContinue
    if ($launcher) {
        $previousErrorPreference = $ErrorActionPreference
        try {
            $ErrorActionPreference = 'Continue'
            $candidate = & $launcher.Source -3 -c 'import sys; print(sys.executable)' 2>$null
            $pythonFound = $LASTEXITCODE -eq 0
        } finally {
            $ErrorActionPreference = $previousErrorPreference
        }
        if ($pythonFound -and $candidate -and (Test-Path -LiteralPath $candidate)) { return $candidate }
    }
    $bundles = Get-ChildItem -Path "$env:ProgramFiles\WindowsApps\OpenAI.CodexPrimaryRuntime*\dependencies\python\python.exe" -File -ErrorAction SilentlyContinue | Sort-Object FullName -Descending
    if ($bundles) { return $bundles[0].FullName }
    throw 'Python 3.11+ is required. Install Python from python.org, then run install.bat.'
}

function Install-AppPackages([string]$BootstrapPython) {
    $appPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
    if (-not (Test-Path -LiteralPath $appPython)) {
        Write-Host 'Preparing the application environment...'
        & $BootstrapPython -m venv (Join-Path $PSScriptRoot '.venv')
        if ($LASTEXITCODE -ne 0) { throw 'Could not create the Python virtual environment. Check the folder permissions and Python installation.' }
    }
    & $appPython -m pip install --disable-pip-version-check -r (Join-Path $PSScriptRoot 'requirements.txt')
    if ($LASTEXITCODE -ne 0) { throw 'Package installation failed. Check the internet connection, then run install.bat again.' }
    & $appPython -m approval_generator.dependencies
    if ($LASTEXITCODE -ne 0) { throw 'Required components are still unavailable. See the message above and run install.bat again.' }
}

try {
    $runtimePython = Find-Python
    if ($Install) {
        Install-AppPackages $runtimePython
        Write-Host 'Installation complete. Run start.bat.'
    } else {
        & $runtimePython -m approval_generator.dependencies
        $dependencyStatus = $LASTEXITCODE
        if ($dependencyStatus -eq 2) { throw 'Please repair the Python installation as described above.' }
        if ($dependencyStatus -ne 0) {
            Write-Host 'Installing required components, including XLS support. Internet access is required.'
            Write-Host 'Please wait. The application will open when installation finishes.'
            Install-AppPackages $runtimePython
            $runtimePython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
        }
        & $runtimePython (Join-Path $PSScriptRoot 'main.py')
        if ($LASTEXITCODE -ne 0) { throw 'The application stopped with an error. See the message above.' }
    }
} catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
