$ErrorActionPreference = 'Stop'
$oldCache = $env:PYINSTALLER_CONFIG_DIR
Push-Location -LiteralPath $PSScriptRoot
try {
    $buildEnv = '.venv-build-pyqt6'
    $buildPython = Join-Path $buildEnv 'Scripts\python.exe'
    if (-not (Test-Path -LiteralPath $buildPython)) {
        python -m venv $buildEnv
        if ($LASTEXITCODE) { throw 'Could not create the build environment.' }
    }
    if (-not (Select-String -Quiet -LiteralPath (Join-Path $buildEnv 'pyvenv.cfg') -Pattern '^include-system-site-packages\s*=\s*false\s*$')) {
        throw 'The build environment must be isolated from system site-packages.'
    }
    & $buildPython -I -m pip install --disable-pip-version-check --only-binary=:all: -r requirements.txt -c requirements-build.txt PyInstaller
    if ($LASTEXITCODE) { throw 'Build dependency installation failed.' }
    & $buildPython -I -m pip check
    if ($LASTEXITCODE) { throw 'Incompatible build dependencies.' }
    $env:PYINSTALLER_CONFIG_DIR = Join-Path $PSScriptRoot 'build\pyinstaller-cache'
    & $buildPython -I -m PyInstaller --noconfirm --clean workbench.spec
    if ($LASTEXITCODE) { throw 'Packaging failed.' }
    Get-ChildItem -LiteralPath 'dist' -Filter '*.exe' | Select-Object Name, Length, FullName
} finally {
    $env:PYINSTALLER_CONFIG_DIR = $oldCache
    Pop-Location
}
