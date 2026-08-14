[CmdletBinding()]
param(
    [switch] $Open
)

$ErrorActionPreference = "Stop"

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$doxyfilePath = Join-Path $repositoryRoot "Doxyfile"
$documentationDirectory = Join-Path $repositoryRoot "Doxygen_Documentation_HTML"
$documentationIndex = Join-Path $repositoryRoot "Doxygen_Documentation_HTML\html\index.html"
$documentationHtmlDirectory = Split-Path -Parent $documentationIndex
$warningsLog = Join-Path $documentationDirectory "doxygen-warnings.log"
$localToolsDirectory = Join-Path $repositoryRoot ".tools"

function Set-DarkThemeForGeneratedGraphs {
    param(
        [Parameter(Mandatory = $true)] [string] $HtmlDirectory
    )

    # Doxygen жёстко записывает светлую палитру class/collaboration/directory-графов в SVG,
    # поэтому обычный HTML_COLORSTYLE и внешний CSS на содержимое iframe не влияют.
    $darkGraphStyle = @'
<style type="text/css"><![CDATA[
svg { background: #1b1f26; }
.graph > polygon { fill: #1b1f26; }
.node polygon:not([fill="none"]):not([fill="#666666"]),
.node ellipse:not([fill="none"]):not([fill="#666666"]) { fill: #273142; stroke: #718096; }
.node polygon[fill="none"], .node ellipse[fill="none"] { fill: none; stroke: #718096; }
.node polygon[fill="#999999"], .node ellipse[fill="#999999"] { fill: #253b56; }
.node text { fill: #f1f5f9; }
.edge path, .edge polyline { stroke: #8aa4c2; }
.edge polygon { fill: #8aa4c2; stroke: #8aa4c2; }
.edge text { fill: #d9e2ec; }
]]></style>
'@
    $svgRootPattern = [regex]::new('<svg\b[^>]*>', [Text.RegularExpressions.RegexOptions]::Singleline)
    $automaticGraphPattern = '(__graph|_dep)\.svg$|^inherit_graph_\d+\.svg$|^graph_legend\.svg$'

    $graphFiles = Get-ChildItem -LiteralPath $HtmlDirectory -Filter "*.svg" -File |
        Where-Object { $_.Name -match $automaticGraphPattern }

    foreach ($graphFile in $graphFiles) {
        $svg = Get-Content -LiteralPath $graphFile.FullName -Raw
        if (!$svgRootPattern.IsMatch($svg)) {
            throw "Generated graph does not contain an SVG root: $($graphFile.FullName)"
        }

        $styledSvg = $svgRootPattern.Replace(
            $svg,
            { param($match) $match.Value + [Environment]::NewLine + $darkGraphStyle },
            1
        )

        [IO.File]::WriteAllText($graphFile.FullName, $styledSvg, [Text.UTF8Encoding]::new($false))
    }

    Write-Host "Dark theme applied to generated automatic graphs: $($graphFiles.Count)"
}

$localDoxygen = Get-ChildItem -LiteralPath (Join-Path $localToolsDirectory "doxygen") -Filter "doxygen.exe" -Recurse -File -ErrorAction SilentlyContinue |
    Sort-Object FullName -Descending |
    Select-Object -First 1
$doxygenExecutable = if ($null -ne $localDoxygen) {
    $localDoxygen.FullName
}
else {
    $pathDoxygen = Get-Command doxygen -ErrorAction SilentlyContinue
    if ($null -ne $pathDoxygen) {
        $pathDoxygen.Source
    }
}

if ([string]::IsNullOrEmpty($doxygenExecutable)) {
    throw @"
Doxygen was not found in .tools or PATH.
Run docs\BootstrapDocumentationTools.ps1 or install the Windows binary from
https://www.doxygen.nl/download.html and retry.
"@
}

$localDot = Get-ChildItem -LiteralPath (Join-Path $localToolsDirectory "graphviz") -Filter "dot.exe" -Recurse -File -ErrorAction SilentlyContinue |
    Sort-Object FullName -Descending |
    Select-Object -First 1
$dotExecutable = if ($null -ne $localDot) {
    $localDot.FullName
}
else {
    $pathDot = Get-Command dot -ErrorAction SilentlyContinue
    if ($null -ne $pathDot) {
        $pathDot.Source
    }
}

if ([string]::IsNullOrEmpty($dotExecutable)) {
    throw @"
Graphviz dot was not found in .tools or PATH.
Run docs\BootstrapDocumentationTools.ps1 or install Graphviz from
https://graphviz.org/download/ and retry. Doxygen needs dot for all diagrams.
"@
}

$dotDirectory = Split-Path -Parent $dotExecutable
$env:PATH = "$dotDirectory;$env:PATH"

if (Test-Path -LiteralPath $documentationDirectory) {
    $resolvedRepositoryRoot = [IO.Path]::GetFullPath($repositoryRoot).TrimEnd('\') + '\'
    $resolvedDocumentationDirectory = [IO.Path]::GetFullPath($documentationDirectory).TrimEnd('\') + '\'
    if (!$resolvedDocumentationDirectory.StartsWith($resolvedRepositoryRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clean documentation outside the repository: $resolvedDocumentationDirectory"
    }

    Remove-Item -LiteralPath $documentationDirectory -Recurse -Force
}

Push-Location $repositoryRoot
try {
    & $doxygenExecutable $doxyfilePath
    if ($LASTEXITCODE -ne 0) {
        throw "Doxygen exited with code $LASTEXITCODE."
    }
}
finally {
    Pop-Location
}

if (!(Test-Path -LiteralPath $documentationIndex)) {
    throw "Doxygen did not create the expected index: $documentationIndex"
}

Set-DarkThemeForGeneratedGraphs -HtmlDirectory $documentationHtmlDirectory

Write-Host "Documentation generated: $documentationIndex"
if (Test-Path -LiteralPath $warningsLog) {
    $warningCount = @(Get-Content -LiteralPath $warningsLog).Count
    Write-Host "Doxygen warnings: $warningCount ($warningsLog)"
}

if ($Open) {
    Start-Process $documentationIndex
}
