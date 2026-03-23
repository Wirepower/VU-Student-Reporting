<#
.SYNOPSIS
  Publishes the WinForms app for packaging into an MSI (or other installer).

.DESCRIPTION
  - Framework-dependent (default): smaller output; target PCs need .NET 8 Desktop Runtime installed.
  - Self-contained: larger output; includes .NET runtime — easier for "install everything" MSI payloads.

  Run from repo root:
    .\build\Publish-ForInstaller.ps1
    .\build\Publish-ForInstaller.ps1 -SelfContained

.NOTES
  Office (Outlook/Excel) and VPN/P-drive assumptions are documented in docs/MSI-Installer-Guide.md — they are not bundled.
#>

param(
    [switch]$SelfContained,
    [string]$Runtime = "win-x64",
    [string]$Configuration = "Release",
    [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $repoRoot

$projectFile = Join-Path $repoRoot "Student Attendance Reporting.vbproj"
if (-not (Test-Path $projectFile)) {
    throw "Project not found: $projectFile"
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $repoRoot "publish\installer-payload\$Runtime"
}

Write-Host "Repo: $repoRoot"
Write-Host "Output: $OutputPath"
Write-Host "Configuration: $Configuration | Runtime: $Runtime | SelfContained: $SelfContained"

$args = @(
    "publish", "`"$projectFile`"",
    "-c", $Configuration,
    "-r", $Runtime,
    "-o", "`"$OutputPath`""
)

if ($SelfContained) {
    $args += "--self-contained", "true"
} else {
    $args += "--self-contained", "false"
}

# Trim unused locale satellites in self-contained builds (optional size win)
if ($SelfContained) {
    $args += "-p:PublishTrimmed=false"
}

Write-Host "dotnet $($args -join ' ')"
& dotnet @args

if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed with exit code $LASTEXITCODE"
}

Write-Host ""
Write-Host "Done. Payload folder ready for your MSI/Setup project:"
Write-Host "  $OutputPath"
Write-Host ""
Write-Host "Next: add this folder's contents to your Visual Studio Setup Project (see docs/MSI-Installer-Guide.md)."
