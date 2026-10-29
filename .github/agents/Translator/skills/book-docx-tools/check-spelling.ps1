# Proofing aid for the ORIGINAL-language Word file (phase 1). Read-only (works on a temporary copy).
# Collects Word spelling errors per block (in the block's language) and regex suspects typical of OCR, and writes
# proof-sheet queries so every suspect word can be checked against the page image with pdf-source-tools\proof-sheet.mjs.
# Usage: & check-spelling.ps1 -Project project.json -Out suspects.json [-Queries queries.json] [-Allow allow.txt] [-Docx other.docx]
param([Parameter(Mandatory)][string]$Project, [Parameter(Mandatory)][string]$Out, [string]$Queries = '', [string]$Allow = '', [string]$Docx = '')
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'BookDocx.psm1') -Force
$p = Read-BookProject $Project
$blocks = Get-BookBlocks $p
$allow = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
if ($Allow) { Get-Content -LiteralPath $Allow -Encoding utf8 | ForEach-Object { $w = $_.Trim(); if ($w -and -not $w.StartsWith('#')) { [void]$allow.Add($w) } } }
$patterns = [ordered]@{
    pipe          = '\|'
    double_single = "‘‘|’’|''|,,"
    digit_in_word = '\p{L}\d+\p{L}|\b[Il1]\d{0,1}[a-z]{2,}\d'
    ligature      = '[\uFB00-\uFB06]'
    split_hyphen  = '\p{Ll}- \p{Ll}'
    double_space  = '\S {2,}\S'
    odd_char      = '[^\p{L}\p{N}\p{P}\p{Zs}\p{Sm}\p{Sc}\p{Sk}\u00B0\u00B1\u00B2\u00B3\u00B5\u00B7\u00D7\u2032\u2033\n\t]'
}
$target = if ($Docx) { (Resolve-Path -LiteralPath $Docx).Path } else { $p.Output }
$temp = Join-Path ([IO.Path]::GetTempPath()) ("book-spell-" + [guid]::NewGuid().ToString('N') + '.docx')
Copy-Item -LiteralPath $target -Destination $temp
$items = New-Object System.Collections.Generic.List[object]
$word = New-OwnedWord
$doc = $null
try {
    $doc = $word.Documents.Open($temp, $false, $true, $false)
    for ($i = 1; $i -le $blocks.Count; $i++) {
        $b = $blocks[$i - 1]; $name = Get-BlockBookmark $p $i
        if (-not $doc.Bookmarks.Exists($name)) { continue }
        $pages = @($b.source_pages | ForEach-Object { [int]$_ })
        if ($pages.Count -eq 2 -and $pages[1] -gt $pages[0] + 1) { $pages = @($pages[0]..$pages[1]) }
        $range = $doc.Bookmarks.Item($name).Range
        $isTable = $b.type -eq 'table'
        $text = if ($isTable) { (@($b.rows | ForEach-Object { @($_) -join "`t" }) -join "`n") } else { [string]$b.text }
        $n = 0
        foreach ($e in $range.SpellingErrors) {
            $w = ([string]$e.Text).Trim()
            if (-not $w -or $allow.Contains($w)) { continue }
            $n++
            $cs = if ($isTable) { 0 } else { [Math]::Min($text.Length, [Math]::Max(0, $e.Start - $range.Start - 40)) }
            $items.Add([pscustomobject]@{ id = "$name-s$n"; block = $name; kind = 'spelling'; word = $w; pages = $pages; context = $text.Substring($cs, [Math]::Min(100, $text.Length - $cs)) })
        }
        foreach ($k in $patterns.Keys) {
            foreach ($m in [regex]::Matches($text, $patterns[$k])) {
                $n++
                $s = [Math]::Max(0, $m.Index - 40)
                $token = ([regex]::Match($text.Substring($s), '\S*' + [regex]::Escape($m.Value.Trim()) + '\S*')).Value
                if (-not $token) { $token = $m.Value }
                if ($allow.Contains($token.Trim('.,;:()"“”‘’'))) { continue }
                $items.Add([pscustomobject]@{ id = "$name-r$n"; block = $name; kind = $k; word = $token; pages = $pages; context = $text.Substring($s, [Math]::Min(100, $text.Length - $s)) })
            }
        }
    }
    $doc.Close(0); $doc = $null
} finally { Close-OwnedWord $word $doc; Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
$items | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $Out -Encoding utf8
if ($Queries) {
    @($items | ForEach-Object {
        $w = ($_.word -replace '^[^\p{L}\p{N}]+|[^\p{L}\p{N}]+$', '')
        if ($w) { [pscustomobject]@{ id = $_.id; page = $_.pages[0]; pages = $_.pages; word = $w; note = $_.kind } }
    }) | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $Queries -Encoding utf8
}
$items | Group-Object kind | ForEach-Object { Write-Output ("{0,-14} {1}" -f $_.Name, $_.Count) }
Write-Output "Total suspects: $($items.Count) -> $Out"
