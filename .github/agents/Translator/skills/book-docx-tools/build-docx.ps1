# Builds a NEW editable book DOCX from a project config and its block batches. Never overwrites.
# Usage: & build-docx.ps1 -Project project.json [-Preview kontrola.pdf]
param([Parameter(Mandatory)][string]$Project, [string]$Preview = '')
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'BookDocx.psm1') -Force
$p = Read-BookProject $Project
$blocks = Get-BookBlocks $p
[void](Test-SourceUnchanged $p)
$working = Join-Path $p.Dir ([IO.Path]::GetFileNameWithoutExtension($p.ConfigPath) + '-working.docx')
foreach ($f in @($p.Output, $working)) { if (Test-Path -LiteralPath $f) { throw "File exists; preserve it: $f" } }
if ($Preview -and (Test-Path -LiteralPath $Preview)) { throw "Preview exists: $Preview" }
$word = New-OwnedWord
$doc = $null
try {
    $doc = $word.Documents.Add()
    Initialize-BookDocument $doc $word $p
    if ($p.TitlePage) { Add-BookTitlePages $doc $p }
    [void](Add-BookToc $doc $p)
    for ($i = 1; $i -le $blocks.Count; $i++) {
        Add-BookBlock $doc $word $p $blocks[$i - 1] (Get-BlockBookmark $p $i)
        if ($i % 200 -eq 0) { Write-Host "  $i / $($blocks.Count) blocks" }
    }
    Set-BookFooter $doc
    [void](Update-BookToc $doc $p)
    $doc.SaveAs2($working, 16)
    $doc.Close(0); $doc = $null
    $doc = $word.Documents.Open($working, $false, $true, $false)
    $doc.Repaginate()
    $check = Test-BookDocument $doc $p $blocks
    if ($Preview) { $doc.ExportAsFixedFormat([IO.Path]::GetFullPath($Preview), 17) }
    $doc.Close(0); $doc = $null
} finally { Close-OwnedWord $word $doc }
$package = Test-DocxPackage $working
if ($check.problems.Count -or $check.text_differences.Count) {
    $check | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath ($working + '.problems.json') -Encoding utf8
    throw "Verification failed ($($check.problems.Count) problems, $($check.text_differences.Count) text differences). Working copy kept: $working"
}
[void](Test-SourceUnchanged $p)
if (Test-Path -LiteralPath $p.Output) { throw 'Output appeared during the build; refusing to overwrite.' }
Move-Item -LiteralPath $working -Destination $p.Output
$report = [ordered]@{
    action = 'build'; time = (Get-Date).ToString('s')
    output = $p.Output; output_sha256 = (Get-FileHash -LiteralPath $p.Output -Algorithm SHA256).Hash
    source = $p.Source; source_sha256 = $p.SourceSha256; source_unchanged = [bool]$p.SourceSha256
    batches = @($p.Batches | ForEach-Object { [IO.Path]::GetFileName($_) })
    package = $package
}
foreach ($prop in $check.PSObject.Properties) { $report[$prop.Name] = $prop.Value }
Write-BookReport $p $report
Write-Output ("Built {0}: {1} blocks, {2} table cells, {3} TOC links, {4} pages, {5} language warnings." -f $p.Output, $check.blocks_checked, $check.table_cells_checked, $check.toc_links.Count, $check.document_pages, $check.language_warnings)
