$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

# Apply idempotent browser-script fixes before validation/build.
python .\patch_auto_update_js.py
if ($LASTEXITCODE -ne 0) { throw "patch_auto_update_js.py failed with exit code $LASTEXITCODE" }

$distDir = Join-Path $PSScriptRoot 'dist'
$buildDir = Join-Path $PSScriptRoot 'build'
$runtimeDir = Join-Path $PSScriptRoot 'KafalaCompareApp_build'
$resolvedRoot = (Resolve-Path $PSScriptRoot).Path

function Assert-InProject([string]$PathToCheck) {
    if (-not (Test-Path $PathToCheck)) { return }
    $resolved = (Resolve-Path $PathToCheck).Path
    if (-not $resolved.StartsWith($resolvedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to modify outside project folder: $resolved"
    }
}

foreach ($dir in @($distDir, $buildDir)) {
    if (Test-Path $dir) {
        Assert-InProject $dir
        Remove-Item -LiteralPath $dir -Recurse -Force
    }
}

# Prepare the Node.js runtime that PyInstaller bundles inside the EXE.
# This makes the repository buildable on a clean Windows machine / GitHub Actions.
if (Test-Path $runtimeDir) {
    Assert-InProject $runtimeDir
    Remove-Item -LiteralPath $runtimeDir -Recurse -Force
}
New-Item -ItemType Directory -Path $runtimeDir | Out-Null

$nodeCommand = Get-Command node.exe -ErrorAction SilentlyContinue
if (-not $nodeCommand) {
    $nodeCommand = Get-Command node -ErrorAction SilentlyContinue
}
if (-not $nodeCommand) {
    throw 'Node.js was not found. Install Node.js 22 or run the GitHub Actions build workflow.'
}

& $nodeCommand.Source --check (Join-Path $PSScriptRoot 'auto_update_from_diff.js')
if ($LASTEXITCODE -ne 0) { throw "Patched auto_update_from_diff.js failed syntax validation" }

Copy-Item -LiteralPath $nodeCommand.Source -Destination (Join-Path $runtimeDir 'node.exe') -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'package.json') -Destination $runtimeDir -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'package-lock.json') -Destination $runtimeDir -Force

$oldSkip = $env:PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD
$env:PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD = '1'
try {
    Push-Location $runtimeDir
    npm ci --omit=dev
    if ($LASTEXITCODE -ne 0) { throw "npm ci failed with exit code $LASTEXITCODE" }
    Pop-Location
}
finally {
    if ((Get-Location).Path -eq $runtimeDir) { Pop-Location }
    $env:PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD = $oldSkip
}

if (-not (Test-Path (Join-Path $runtimeDir 'node_modules\playwright'))) {
    throw 'Playwright dependency was not prepared correctly.'
}
if (-not (Test-Path (Join-Path $runtimeDir 'node_modules\xlsx'))) {
    throw 'xlsx dependency was not prepared correctly.'
}

# Build the one-file Windows application using the active Python from PATH.
python -m PyInstaller --clean --noconfirm KafalaCompareApp_fixed.spec
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE" }

$exePath = Join-Path $distDir 'KafalaCompareApp.exe'
if (-not (Test-Path $exePath)) {
    throw "Expected EXE was not created: $exePath"
}

# Create portable ZIP from the clean output.
$zipPath = Join-Path $PSScriptRoot 'KafalaCompareApp_build.zip'
if (Test-Path $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}
Compress-Archive -Path $exePath -DestinationPath $zipPath -Force

Write-Host "EXE built at: $exePath"
Write-Host "Portable package built at: $zipPath"

# Download VC++ Redistributable for the installer.
$vcRedistUrl = 'https://aka.ms/vs/17/release/vc_redist.x64.exe'
$vcRedistPath = Join-Path $PSScriptRoot 'vc_redist.x64.exe'
if (-not (Test-Path $vcRedistPath)) {
    Write-Host 'Downloading Visual C++ Redistributable...'
    Invoke-WebRequest -Uri $vcRedistUrl -OutFile $vcRedistPath -UseBasicParsing
}

# Compile Inno Setup installer when the compiler is available.
$isccCandidates = @(
    'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
    'C:\Program Files\Inno Setup 6\ISCC.exe'
)
$isccPath = $isccCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if ($isccPath) {
    Write-Host 'Compiling Inno Setup installer...'
    & $isccPath '.\KafalaCompareApp.iss'
    if ($LASTEXITCODE -ne 0) { throw "Inno Setup failed with exit code $LASTEXITCODE" }
    Write-Host 'Installer created successfully.'
} else {
    Write-Host 'Inno Setup 6 not found; portable EXE/ZIP were still built successfully.'
}

# PyInstaller temporary build directory is safe to remove. Keep runtimeDir because
# it can help diagnose packaging issues on the developer machine.
if (Test-Path $buildDir) {
    Assert-InProject $buildDir
    Remove-Item -LiteralPath $buildDir -Recurse -Force
    Write-Host 'Temporary PyInstaller build folder removed.'
}
