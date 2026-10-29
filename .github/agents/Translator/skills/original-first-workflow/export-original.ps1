# Phase 2 entry point: reads the CURRENT checked original .docx (including the user's later edits) block by block and
# writes (1) original-blocks.json and (2) translation batch skeletons with "src" links, the original text in "orig"
# and an empty "text" to be translated. The source PDF is not needed any more.
# Usage: & export-original.ps1 -Project original.json -OutDir prace [-Prefix cz] [-BatchSize 40] [-Force]
param([Parameter(Mandatory)][string]$Project, [Parameter(Mandatory)][string]$OutDir, [string]$Prefix = 'translation', [int]$BatchSize = 40, [switch]$Force)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..\book-docx-tools\BookDocx.psm1') -Force
$p = Read-BookProject $Project
$blocks = Get-BookBlocks $p
$docx = $p.Output
$hash = (Get-FileHash -LiteralPath $docx -Algorithm SHA256).Hash
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$OutDir = (Resolve-Path -LiteralPath $OutDir).Path
$allPath = Join-Path $OutDir 'original-blocks.json'
if ((Test-Path -LiteralPath $allPath) -and -not $Force) { throw "$allPath exists; use -Force to refresh it (skeleton batches are never overwritten)." }

$temp = Join-Path ([IO.Path]::GetTempPath()) ('book-export-' + [guid]::NewGuid().ToString('N') + '.docx')
Copy-Item -LiteralPath $docx -Destination $temp
$word = New-OwnedWord; $doc = $null
$out = New-Object System.Collections.Generic.List[object]
$warnings = New-Object System.Collections.Generic.List[string]
try {
    $doc = $word.Documents.Open($temp, $false, $true, $false)
    $prevEnd = $null; $prevName = '(start)'
    if ($doc.TablesOfContents.Count -ge 1) { $prevEnd = $doc.TablesOfContents.Item(1).Range.End; $prevName = '(TOC)' }
    for ($i = 1; $i -le $blocks.Count; $i++) {
        $b = $blocks[$i - 1]; $name = Get-BlockBookmark $p $i
        if (-not $doc.Bookmarks.Exists($name)) { throw "Missing block bookmark $name in $docx; audit the original first." }
        $range = $doc.Bookmarks.Item($name).Range
        if ($null -ne $prevEnd -and $range.Start -gt $prevEnd) {
            $gap = $doc.Range($prevEnd, $range.Start).Text -replace "[\r\n\v\f\a/]", '' -replace '\s+', ''
            if ($gap) { $warnings.Add("Text without a block bookmark between $prevName and ${name}: $($doc.Range($prevEnd, $range.Start).Text.Trim())") }
        }
        $item = [ordered]@{ src = $name; type = $b.type }
        if ($b.type -eq 'table') {
            $table = $range.Tables.Item(1)
            $rows = @(for ($r = 1; $r -le $table.Rows.Count; $r++) { , @(for ($c = 1; $c -le $table.Columns.Count; $c++) { $table.Cell($r, $c).Range.Text.TrimEnd([char[]]@([char]13, [char]7)) }) })
            $item.rows = $rows
            $prevEnd = $table.Range.End
        } else {
            $item.text = $range.Text.TrimEnd([char]13).Replace("`v", "`n")
            $prevEnd = $range.Paragraphs.Item($range.Paragraphs.Count).Range.End
        }
        $prevName = $name
        foreach ($k in @('file', 'width_cm', 'col_widths_cm', 'source_pages', 'new_page', 'lang')) { if ($null -ne $b.$k) { $item[$k] = $b.$k } }
        $out.Add([pscustomobject]$item)
    }
    $tail = $doc.Range($prevEnd, $doc.Content.End).Text -replace "[\r\n\v\f\a/]", '' -replace '\s+', ''
    if ($tail) { $warnings.Add("Text without a block bookmark after ${prevName}: $($doc.Range($prevEnd, $doc.Content.End).Text.Trim())") }
    $doc.Close(0); $doc = $null
} finally { Close-OwnedWord $word $doc; Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
if ((Get-FileHash -LiteralPath $docx -Algorithm SHA256).Hash -ne $hash) { throw 'The original changed during export; repeat it.' }

[ordered]@{ source_docx = $docx; source_sha256 = $hash; assets_dir = $p.AssetsDir; time = (Get-Date).ToString('s'); warnings = $warnings; blocks = $out } |
    ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $allPath -Encoding utf8

# Skeleton batches: split at h1/h2 once the batch reaches BatchSize blocks (or at 2*BatchSize regardless).
$batches = New-Object System.Collections.Generic.List[object]; $cur = New-Object System.Collections.Generic.List[object]
foreach ($o in $out) {
    if ($cur.Count -ge 2 * $BatchSize -or ($cur.Count -ge $BatchSize -and $o.type -in @('h1', 'h2'))) { $batches.Add($cur); $cur = New-Object System.Collections.Generic.List[object] }
    $cur.Add($o)
}
if ($cur.Count) { $batches.Add($cur) }
$n = 0; $written = 0
foreach ($batch in $batches) {
    $n++
    $path = Join-Path $OutDir ('{0}-{1:D2}.json' -f $Prefix, $n)
    if (Test-Path -LiteralPath $path) { continue }
    $pages = @($batch | ForEach-Object { $_.source_pages } | Where-Object { $null -ne $_ } | ForEach-Object { [int]$_ })
    $tb = foreach ($o in $batch) {
        $t = [ordered]@{ src = $o.src; type = $o.type }
        if ($o.type -eq 'table') { $t.orig_rows = $o.rows; $t.rows = @($o.rows | ForEach-Object { , @($_ | ForEach-Object { '' }) }) }
        else { $t.orig = $o.text; $t.text = '' }
        foreach ($k in @('file', 'width_cm', 'col_widths_cm', 'source_pages', 'new_page')) { if ($o.PSObject.Properties[$k]) { $t[$k] = $o.$k } }
        [pscustomobject]$t
    }
    [ordered]@{ source_sha256 = $hash; source_pages = $(if ($pages.Count) { @([int]($pages | Measure-Object -Minimum).Minimum, [int]($pages | Measure-Object -Maximum).Maximum) } else { @() }); status = 'todo'; blocks = @($tb) } |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $path -Encoding utf8
    $written++
}
Write-Output ("Exported {0} blocks from {1} (SHA256 {2}) -> {3}; {4} skeleton batches written, {5} already existed. Warnings: {6}" -f $out.Count, $docx, $hash, $allPath, $written, ($batches.Count - $written), $warnings.Count)
Write-Output ("Translation project: source_file = the original .docx above, assets_dir = {0} (image 'file' values are relative to it)." -f $p.AssetsDir)
foreach ($w in $warnings) { Write-Output "  WARNING: $w" }
