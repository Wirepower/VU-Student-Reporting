<#
.SYNOPSIS
  Builds an MSI from publish output using WiX Toolset.

.DESCRIPTION
  - Harvests all files from a publish directory (Heat)
  - Compiles WiX source to objects (Candle)
  - Links objects into a single MSI (Light)

  This script is intended for CI/CD and local scripted builds.
#>

param(
    [Alias("PublishDir")]
    [string]$PayloadDirectory = "",
    [Alias("OutputDir")]
    [string]$OutputDirectory = "",
    [string]$ProductName = "Student Attendance Reporting",
    [string]$Manufacturer = "Victoria University",
    [string]$ProductVersion = "1.0.0",
    [string]$MsiName = "StudentAttendanceReporting-Setup.msi"
)

$ErrorActionPreference = "Stop"

Set-StrictMode -Version Latest
$repoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $repoRoot

if ([string]::IsNullOrWhiteSpace($PayloadDirectory)) {
    $PayloadDirectory = Join-Path $repoRoot "publish\installer-payload\win-x64"
}
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $repoRoot "artifacts"
}

function Get-WixToolPath {
    param([string]$ToolName)

    $cmd = Get-Command $ToolName -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    $commonPaths = @(
        "C:\Program Files (x86)\WiX Toolset v3.11\bin\$ToolName",
        "C:\Program Files\WiX Toolset v3.11\bin\$ToolName"
    )

    foreach ($path in $commonPaths) {
        if (Test-Path $path) {
            return $path
        }
    }

    throw "Unable to locate $ToolName. Install WiX Toolset v3.11 and ensure it is on PATH."
}

if (-not (Test-Path $PayloadDirectory)) {
    throw "Publish directory not found: $PayloadDirectory"
}

$wixSourceDir = Join-Path $repoRoot "installer\wix"
$productWxs = Join-Path $wixSourceDir "Product.wxs"
$wixObjDir = Join-Path $OutputDirectory "wixobj"
$harvestWxs = Join-Path $wixObjDir "AppFiles.wxs"
$msiPath = Join-Path $OutputDirectory $MsiName
$shortcutIconPath = Join-Path $repoRoot "VU Support Hub_Desktop Icon-Favicon.ico"

if (-not (Test-Path $productWxs)) {
    throw "WiX source not found: $productWxs"
}
if (-not (Test-Path $shortcutIconPath)) {
    throw "Shortcut icon file not found: $shortcutIconPath"
}

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $wixObjDir -Force | Out-Null

$heat = Get-WixToolPath -ToolName "heat.exe"
$candle = Get-WixToolPath -ToolName "candle.exe"
$light = Get-WixToolPath -ToolName "light.exe"

Write-Host "PublishDir: $PayloadDirectory"
Write-Host "OutputDir:  $OutputDirectory"
Write-Host "MSI:        $msiPath"
Write-Host "Version:    $ProductVersion"

& $heat dir "$PayloadDirectory" `
    -nologo `
    -ag `
    -sreg `
    -scom `
    -sfrag `
    -srd `
    -cg AppFiles `
    -dr INSTALLFOLDER `
    -var var.PublishDir `
    -out "$harvestWxs"

& $candle `
    -nologo `
    -arch x64 `
    -dPublishDir="$PayloadDirectory" `
    -dProductName="$ProductName" `
    -dManufacturer="$Manufacturer" `
    -dProductVersion="$ProductVersion" `
    -dAppIconPath="$shortcutIconPath" `
    -out "$wixObjDir\" `
    "$productWxs" `
    "$harvestWxs"

& $light `
    -nologo `
    -sval `
    -out "$msiPath" `
    (Join-Path $wixObjDir "Product.wixobj") `
    (Join-Path $wixObjDir "AppFiles.wixobj")

if (-not (Test-Path $msiPath)) {
    throw "MSI build did not produce expected file: $msiPath"
}

Write-Host "MSI build complete: $msiPath"
