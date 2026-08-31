$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$distDir = Join-Path $PSScriptRoot "dist"
$buildDir = Join-Path $PSScriptRoot "build"
$resolvedRoot = (Resolve-Path $PSScriptRoot).Path

foreach ($dir in @($distDir, $buildDir)) {
    if (Test-Path $dir) {
        $resolvedDir = (Resolve-Path $dir).Path
        if (-not $resolvedDir.StartsWith($resolvedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to clean outside project folder: $resolvedDir"
        }
        Remove-Item -LiteralPath $resolvedDir -Recurse -Force
    }
}

# Run PyInstaller with the corrected spec file.
py -3 -m PyInstaller --clean --noconfirm KafalaCompareApp_fixed.spec

# Create portable zip from the clean output.
$zipPath = Join-Path $PSScriptRoot "KafalaCompareApp_build.zip"
if (Test-Path $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}
Compress-Archive -Path ".\dist\KafalaCompareApp.exe" -DestinationPath $zipPath -Force

Write-Host "EXE built at: $PSScriptRoot\dist\KafalaCompareApp.exe"
Write-Host "Package built at: $zipPath"

# Download VC++ Redistributable.
$vcRedistUrl = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
$vcRedistPath = ".\vc_redist.x64.exe"
if (-not (Test-Path $vcRedistPath)) {
    Write-Host "Downloading Visual C++ Redistributable..."
    Invoke-WebRequest -Uri $vcRedistUrl -OutFile $vcRedistPath -UseBasicParsing
}

# Compile Inno Setup script.
$isccPath = "C:\Program Files (x86)\Inno Setup 6\iscc.exe"
if (Test-Path $isccPath) {
    Write-Host "Compiling Inno Setup script..."
    & $isccPath ".\KafalaCompareApp.iss"
    Write-Host "Installer created successfully."
} else {
    Write-Host "Inno Setup compiler not found at $isccPath"
    Write-Host "Please install Inno Setup 6 to create the installer."
}

if (Test-Path $buildDir) {
    $resolvedBuildDir = (Resolve-Path $buildDir).Path
    if (-not $resolvedBuildDir.StartsWith($resolvedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clean outside project folder: $resolvedBuildDir"
    }
    Remove-Item -LiteralPath $resolvedBuildDir -Recurse -Force
    Write-Host "Temporary build folder removed."
}
