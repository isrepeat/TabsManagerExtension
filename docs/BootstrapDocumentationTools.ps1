[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$doxygenVersion = "1.18.0"
$doxygenSha256 = "e84f54cfd49ef06b0b16536056dbec0c496323de28abcce53a4269463de35eaf"
$doxygenUrl = "https://www.doxygen.nl/files/doxygen-$doxygenVersion.windows.x64.bin.zip"

$graphvizVersion = "15.1.0"
$graphvizSha256 = "c3ee71ff81ab97352082225574a140f20f5d6929d5f33d1097a1fe0e4161962a"
$graphvizUrl = "https://gitlab.com/api/v4/projects/4207231/packages/generic/graphviz-releases/$graphvizVersion/windows_10_cmake_Release_Graphviz-$graphvizVersion-win64.zip"

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$toolsDirectory = Join-Path $repositoryRoot ".tools"
$downloadsDirectory = Join-Path $toolsDirectory "downloads"
$doxygenDirectory = Join-Path $toolsDirectory "doxygen\$doxygenVersion"
$graphvizDirectory = Join-Path $toolsDirectory "graphviz\$graphvizVersion"
$doxygenArchive = Join-Path $downloadsDirectory "doxygen-$doxygenVersion.zip"
$graphvizArchive = Join-Path $downloadsDirectory "graphviz-$graphvizVersion.zip"

function Install-PortableTool {
    param(
        [Parameter(Mandatory = $true)] [string] $Name,
        [Parameter(Mandatory = $true)] [string] $Url,
        [Parameter(Mandatory = $true)] [string] $ExpectedSha256,
        [Parameter(Mandatory = $true)] [string] $ArchivePath,
        [Parameter(Mandatory = $true)] [string] $Destination,
        [Parameter(Mandatory = $true)] [string] $ExecutableName
    )

    $existingExecutable = Get-ChildItem -LiteralPath $Destination -Filter $ExecutableName -Recurse -File -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($null -ne $existingExecutable) {
        Write-Host "$Name is already installed: $($existingExecutable.FullName)"
        return
    }

    New-Item -ItemType Directory -Path $downloadsDirectory -Force | Out-Null
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null

    Write-Host "Downloading $Name..."
    Invoke-WebRequest -Uri $Url -OutFile $ArchivePath -UseBasicParsing

    $actualSha256 = (Get-FileHash -LiteralPath $ArchivePath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actualSha256 -ne $ExpectedSha256) {
        throw "$Name SHA-256 mismatch. Expected $ExpectedSha256, got $actualSha256."
    }

    Expand-Archive -LiteralPath $ArchivePath -DestinationPath $Destination -Force
    $installedExecutable = Get-ChildItem -LiteralPath $Destination -Filter $ExecutableName -Recurse -File |
        Select-Object -First 1
    if ($null -eq $installedExecutable) {
        throw "$ExecutableName was not found under $Destination after extracting $Name."
    }

    Write-Host "$Name installed: $($installedExecutable.FullName)"
}

Install-PortableTool `
    -Name "Doxygen $doxygenVersion" `
    -Url $doxygenUrl `
    -ExpectedSha256 $doxygenSha256 `
    -ArchivePath $doxygenArchive `
    -Destination $doxygenDirectory `
    -ExecutableName "doxygen.exe"

Install-PortableTool `
    -Name "Graphviz $graphvizVersion" `
    -Url $graphvizUrl `
    -ExpectedSha256 $graphvizSha256 `
    -ArchivePath $graphvizArchive `
    -Destination $graphvizDirectory `
    -ExecutableName "dot.exe"

Write-Host "Documentation tools are ready. Run docs\GenerateDocumentation.ps1."
