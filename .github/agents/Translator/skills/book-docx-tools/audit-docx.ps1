# Read-only audit of a saved book DOCX against its project config and batches. Works on a temporary copy.
# Usage: & audit-docx.ps1 -Project project.json [-Docx other.docx] [-Report audit.json] [-SkipTitlePages] [-Baseline older.docx]
param([Parameter(Mandatory)][string]$Project, [string]$Docx = '', [string]$Report = '', [switch]$SkipTitlePages, [string]$Baseline = '')
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'BookDocx.psm1') -Force
$p = Read-BookProject $Project
$blocks = Get-BookBlocks $p
$target = if ($Docx) { (Resolve-Path -LiteralPath $Docx).Path } else { $p.Output }
$hash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
$temp = Join-Path ([IO.Path]::GetTempPath()) ("book-audit-" + [guid]::NewGuid().ToString('N') + '.docx')
Copy-Item -LiteralPath $target -Destination $temp
$word = New-OwnedWord
$doc = $null
try {
    $doc = $word.Documents.Open($temp, $false, $true, $false)
    $doc.Repaginate()
    $check = Test-BookDocument $doc $p $blocks -SkipTitlePages:$SkipTitlePages
    $doc.Close(0); $doc = $null
} finally { Close-OwnedWord $word $doc }
try { $package = Test-DocxPackage $temp $(if ($Baseline) { (Resolve-Path -LiteralPath $Baseline).Path } else { '' }) }
finally { Remove-Item -LiteralPath $temp -Force }
if ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -ne $hash) { throw 'The document changed during the audit; repeat it.' }
$source = $null
if ($p.Source -and $p.SourceSha256) { $source = ((Get-FileHash -LiteralPath $p.Source -Algorithm SHA256).Hash -eq $p.SourceSha256) }
$out = [ordered]@{ action = 'audit'; time = (Get-Date).ToString('s'); document = $target; document_sha256 = $hash; source_unchanged = $source; package = $package }
foreach ($prop in $check.PSObject.Properties) { $out[$prop.Name] = $prop.Value }
if ($Report) { $out | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Report -Encoding utf8 }
Write-Output ("Audit {0}: {1} blocks, {2} table cells, {3} images, {4} TOC links, {5} pages, {6} text differences (user revisions are listed, not fixed), {7} language warnings, {8} problems." -f $target, $check.blocks_checked, $check.table_cells_checked, $check.images_checked, $check.toc_links.Count, $check.document_pages, @($check.text_differences).Count, $check.language_warnings, @($check.problems).Count)
foreach ($x in $check.problems) { Write-Output "  PROBLEM: $x" }
if (@($check.problems).Count -or $source -eq $false) { exit 1 }
