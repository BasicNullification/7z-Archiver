# build-installer.ps1
# Reads AssemblyVersion from ..\My Project\AssemblyInfo.vb
# Then compiles .\inno_setup.iss and passes /DAppVersion=<version>

$ErrorActionPreference = "Stop"

# Paths (relative to this script)
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$assemblyInfoPath = Join-Path $scriptDir "..\My Project\AssemblyInfo.vb"
$issPath          = Join-Path $scriptDir ".\inno_setup.iss"

if (-not (Test-Path $assemblyInfoPath)) {
    throw "AssemblyInfo.vb not found: $assemblyInfoPath"
}
if (-not (Test-Path $issPath)) {
    throw "Inno script not found: $issPath"
}

# Read & extract version
$content = Get-Content -LiteralPath $assemblyInfoPath -Raw

# Matches: <Assembly: AssemblyVersion("1.2.3.4")>
$match = [regex]::Match($content, '<Assembly:\s*AssemblyVersion\("(?<ver>[\d.]+)"\)\s*>', 'IgnoreCase')

if (-not $match.Success) {
    throw "Could not find <Assembly: AssemblyVersion(""a.b.c.d"")> in: $assemblyInfoPath"
}

$version = $match.Groups["ver"].Value
Write-Host "AssemblyVersion found: $version"

# Find ISCC.exe (Inno Setup Compiler)
if ([string]::IsNullOrWhiteSpace($iscc) -or -not (Test-Path -LiteralPath $iscc)) {

    $candidates = @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles}\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles(x86)}\Inno Setup 5\ISCC.exe",
        "${env:ProgramFiles}\Inno Setup 5\ISCC.exe"
    )

    $iscc = @($candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1)

    # If we got a match, unwrap it from the one-item array
    if ($iscc.Count -gt 0) { $iscc = $iscc[0] }
}

if ([string]::IsNullOrWhiteSpace($iscc) -or -not (Test-Path -LiteralPath $iscc)) {
    throw "Could not locate ISCC.exe. Set INNO_ISCC env var to the full path of ISCC.exe."
}

Write-Host "Using ISCC: $iscc"
Write-Host "Compiling: $issPath"
Write-Host "Passing: /DAppVersion=$version"

# Run Inno
$arguments = @(
    "/DAppVersion=$version"
    "$issPath"
)

& $iscc @arguments
$exitCode = $LASTEXITCODE

if ($exitCode -ne 0) {
    throw "ISCC failed with exit code $exitCode"
}


Write-Host "Done. Installer compiled successfully."
