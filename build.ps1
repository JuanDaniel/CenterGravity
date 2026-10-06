#requires -Version 5
<#
.SYNOPSIS
    Builds CenterGravity for every supported Revit release and assembles the
    Autodesk ApplicationPlugins bundle + a distributable zip.

.EXAMPLE
    pwsh ./build.ps1
    pwsh ./build.ps1 -RevitVersions 2026,2027 -Configuration Release
#>
[CmdletBinding()]
param(
    [string[]] $RevitVersions = @('2024', '2025', '2026', '2027'),
    [ValidateSet('Debug', 'Release')]
    [string]   $Configuration = 'Release',
    [string]   $Version       = '2.0.1'
)

$ErrorActionPreference = 'Stop'
$root     = $PSScriptRoot
$project  = Join-Path $root 'CenterGravity\CenterGravity.csproj'
$template = Join-Path $root 'Setup\bundle'
$outRoot  = Join-Path $root 'Setup\Output'
$bundle   = Join-Path $outRoot 'CenterGravity.bundle'

if (Test-Path $bundle) { Remove-Item $bundle -Recurse -Force }
New-Item -ItemType Directory -Force -Path (Join-Path $bundle 'Contents\Resources') | Out-Null

Copy-Item (Join-Path $template 'PackageContents.xml') $bundle -Force
Copy-Item (Join-Path $root 'CenterGravity\Resources\icon.ico') `
          (Join-Path $bundle 'Contents\Resources\icon.ico') -Force

foreach ($rv in $RevitVersions) {
    Write-Host "== Building CenterGravity for Revit $rv ($Configuration) ==" -ForegroundColor Cyan

    dotnet build $project -c $Configuration -p:RevitVersion=$rv -p:Version=$Version --nologo
    if ($LASTEXITCODE -ne 0) { throw "Build failed for Revit $rv" }

    $bin  = Join-Path $root "CenterGravity\bin\$Configuration\$rv"
    $dest = Join-Path $bundle "Contents\$rv"
    New-Item -ItemType Directory -Force -Path $dest | Out-Null

    Copy-Item (Join-Path $bin 'CenterGravity.dll') $dest -Force
    if (Test-Path (Join-Path $bin 'CenterGravity.pdb')) {
        Copy-Item (Join-Path $bin 'CenterGravity.pdb') $dest -Force
    }
    Copy-Item (Join-Path $bin 'Resources') $dest -Recurse -Force
    Copy-Item (Join-Path $template 'CenterGravity.addin') $dest -Force
}

$zip = Join-Path $outRoot "CenterGravity-$Version.zip"
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path $bundle -DestinationPath $zip

Write-Host ''
Write-Host "Bundle : $bundle"  -ForegroundColor Green
Write-Host "Zip    : $zip"     -ForegroundColor Green
Write-Host "Next   : compile Setup\Setup.iss with Inno Setup to produce the installer." -ForegroundColor Green
