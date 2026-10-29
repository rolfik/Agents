# Appends blocks of newly listed batches to the CURRENT (possibly user-revised) DOCX without regenerating it.
# Add the new batch file to "batches" in the project config first. Blocks whose index is greater than the number of
# block bookmarks already in the document are appended at the end; nothing before them is changed.
# Usage: & append-batch.ps1 -Project project.json [-ScopeText "Text oznámení o rozsahu"] [-Preview kontrola.pdf]
param([Parameter(Mandatory)][string]$Project, [string]$ScopeText = '', [string]$Preview = '')
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'BookDocx.psm1') -Force
$p = Read-BookProject $Project
$blocks = Get-BookBlocks $p
[void](Test-SourceUnchanged $p)
if (-not (Test-Path -LiteralPath $p.Output)) { throw "Document not found: $($p.Output). Use build-docx.ps1 for a new document." }
try { $lock = [IO.File]::Open($p.Output, 'Open', 'ReadWrite', 'None'); $lock.Dispose() }
catch { throw "The document is open or locked (close it in Word, or save a copy): $($p.Output)" }
$hash0 = (Get-FileHash -LiteralPath $p.Output -Algorithm SHA256).Hash
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = [IO.Path]::Combine([IO.Path]::GetDirectoryName($p.Output), [IO.Path]::GetFileNameWithoutExtension($p.Output) + ".backup-$stamp.docx")
$working = Join-Path $p.Dir ([IO.Path]::GetFileNameWithoutExtension($p.ConfigPath) + '-working.docx')
if (Test-Path -LiteralPath $working) { throw "Working copy exists from an earlier run; inspect it first: $working" }
Copy-Item -LiteralPath $p.Output -Destination $backup
Copy-Item -LiteralPath $p.Output -Destination $working
$word = New-OwnedWord
$doc = $null
try {
    $doc = $word.Documents.Open($working, $false, $false, $false)
    $tracking = $doc.TrackRevisions
    $doc.TrackRevisions = $false
    $existing = 0
    while ($doc.Bookmarks.Exists((Get-BlockBookmark $p ($existing + 1)))) { $existing++ }
    if ($existing -ge $blocks.Count) { throw "Nothing to append: the document already has $existing of $($blocks.Count) blocks." }
    if ($ScopeText) {
        $sb = "$($p.Prefix)_scope"
        if (-not $doc.Bookmarks.Exists($sb)) { throw "Bookmark $sb not found (title_page.notes item with id 'scope')." }
        $r = $doc.Bookmarks.Item($sb).Range
        $r.Text = $ScopeText
        $doc.Bookmarks.Add($sb, $r) | Out-Null
    }
    $toc = $doc.TablesOfContents.Item(1).Range
    $preText = $doc.Range(0, $toc.Start).Text
    $postText = $doc.Range($toc.End, $doc.Content.End).Text
    $counts = @{ shapes = $doc.InlineShapes.Count; tables = $doc.Tables.Count; comments = $doc.Comments.Count; revisions = $doc.Revisions.Count; footnotes = $doc.Footnotes.Count }
    for ($i = $existing + 1; $i -le $blocks.Count; $i++) { Add-BookBlock $doc $word $p $blocks[$i - 1] (Get-BlockBookmark $p $i) }
    [void](Update-BookToc $doc $p)
    $toc = $doc.TablesOfContents.Item(1).Range
    if ($doc.Range(0, $toc.Start).Text -cne $preText) { throw 'Text before the TOC changed.' }
    $newPost = $doc.Range($toc.End, $doc.Content.End).Text
    if (-not $newPost.StartsWith($postText.Substring(0, $postText.Length - 1), [StringComparison]::Ordinal)) { throw 'Existing text after the TOC changed.' }
    if ($doc.Comments.Count -ne $counts.comments -or $doc.Revisions.Count -ne $counts.revisions -or $doc.Footnotes.Count -ne $counts.footnotes) { throw 'Comments, revisions or footnotes changed.' }
    if ($doc.InlineShapes.Count -lt $counts.shapes -or $doc.Tables.Count -lt $counts.tables) { throw 'Existing images or tables were lost.' }
    $doc.TrackRevisions = $tracking
    $doc.Save()
    $doc.Close(0); $doc = $null
    $doc = $word.Documents.Open($working, $false, $true, $false)
    $doc.Repaginate()
    $check = Test-BookDocument $doc $p $blocks -SkipTitlePages:(-not $p.TitlePage)
    if ($Preview) { $doc.ExportAsFixedFormat([IO.Path]::GetFullPath($Preview), 17) }
    $doc.Close(0); $doc = $null
} catch {
    Close-OwnedWord $word $doc; $word = $null; $doc = $null
    throw "Append failed; the original document was not touched. Backup: $backup. Working copy: $working. $_"
} finally { if ($word) { Close-OwnedWord $word $doc } }
$package = Test-DocxPackage $working $backup
$newBad = @($check.text_differences | Where-Object { [int]($_.block -replace '^.*_b', '') -gt $existing })
$problems = @($check.problems)
if ($newBad.Count -or $problems.Count) {
    $check | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath ($working + '.problems.json') -Encoding utf8
    throw "Verification failed ($($problems.Count) problems, $($newBad.Count) differences in appended blocks). The original was not touched; working copy: $working"
}
if ((Get-FileHash -LiteralPath $p.Output -Algorithm SHA256).Hash -ne $hash0) { throw "The document changed during the append (concurrent edit). Not replacing it. Result kept in $working" }
[IO.File]::Replace($working, $p.Output, [NullString]::Value)
$report = [ordered]@{
    action = 'append'; time = (Get-Date).ToString('s')
    output = $p.Output; output_sha256 = (Get-FileHash -LiteralPath $p.Output -Algorithm SHA256).Hash
    previous_sha256 = $hash0; backup = $backup
    appended_blocks = "$($existing + 1)-$($blocks.Count)"
    user_revisions_in_existing_blocks = @($check.text_differences).Count
    batches = @($p.Batches | ForEach-Object { [IO.Path]::GetFileName($_) })
    package = $package
}
foreach ($prop in $check.PSObject.Properties) { $report[$prop.Name] = $prop.Value }
Write-BookReport $p $report
Write-Output ("Appended blocks {0}-{1} to {2}. {3} TOC links, {4} pages. Existing blocks with user changes: {5}. Backup: {6}" -f ($existing + 1), $blocks.Count, $p.Output, $check.toc_links.Count, $check.document_pages, @($check.text_differences).Count, $backup)
