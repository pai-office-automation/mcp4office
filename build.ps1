<#
.SYNOPSIS
  Compila mcp4office (Release), esegue tutti i self-test (inclusi test profondi JSON/TOML/MSIX/ownership di ClientConfigurator),
  e genera dist\mcp4office-Setup-<version>.exe + SHA256SUMS.txt.
#>
param(
    [string]$Version = ''
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Version)) {
    [xml]$csproj = Get-Content (Join-Path $root 'mcp4office.csproj')
    $Version = $csproj.Project.PropertyGroup.Version | Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($Version)) { throw "Tag <Version> mancante in mcp4office.csproj." }
}

Write-Host "== 1. Compilazione mcp4office ($Version) ==" -ForegroundColor Cyan
dotnet build (Join-Path $root 'mcp4office.csproj') -c Release --nologo
if ($LASTEXITCODE -ne 0) { throw "dotnet build fallito." }

$exe = Join-Path $root 'bin\Release\net48\mcp4office.exe'
if (-not (Test-Path -LiteralPath $exe)) { throw "EXE non trovato: $exe" }

Write-Host "`n== 2. Self-test mcp4office.exe (MCP + Tool + ClientConfigurator JSON/TOML/MSIX) ==" -ForegroundColor Cyan
& $exe --self-test
if ($LASTEXITCODE -ne 0) { throw "Self-test falliti." }

Write-Host "`n== 3. Compilazione Setup Inno Setup ==" -ForegroundColor Cyan
$iscc = @(
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles}\Inno Setup 6\ISCC.exe",
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $iscc) { throw "ISCC.exe (Inno Setup 6) non trovato." }

& $iscc "/DAppVersion=$Version" (Join-Path $root 'installer\mcp4office.iss')
if ($LASTEXITCODE -ne 0) { throw "ISCC fallito." }

$setupExe = Join-Path $root "dist\mcp4office-Setup-$Version.exe"
$hash = (Get-FileHash -LiteralPath $setupExe -Algorithm SHA256).Hash.ToLowerInvariant()
Set-Content -Path (Join-Path $root 'dist\SHA256SUMS.txt') -Value "$hash  mcp4office-Setup-$Version.exe" -Encoding ASCII

Write-Host "`nCOMPLETATO: $setupExe" -ForegroundColor Green
Write-Host "SHA-256:    $hash" -ForegroundColor Green
