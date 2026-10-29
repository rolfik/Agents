# Shared helpers for building, appending and auditing editable book DOCX files through Word COM.
# All paths in a project config are relative to the config file. See SKILL.md and project.example.json.
$ErrorActionPreference = 'Stop'
$script:BlockTypes = @('h1', 'h2', 'h3', 'p', 'note', 'label', 'quote', 'signature', 'image', 'table')
$script:TextStyles = @('Text', 'Note', 'Label', 'Quote', 'Caption', 'Signature', 'Author', 'OriginalTitle', 'Edition')

function Resolve-ProjectPath([string]$Dir, $Value) {
    if (-not $Value) { return $null }
    if ([IO.Path]::IsPathRooted($Value)) { return [IO.Path]::GetFullPath($Value) }
    return [IO.Path]::GetFullPath((Join-Path $Dir $Value))
}

function Read-BookProject([Parameter(Mandatory)][string]$Path) {
    $full = (Resolve-Path -LiteralPath $Path).Path
    $dir = Split-Path $full -Parent
    $c = Get-Content -LiteralPath $full -Raw -Encoding utf8 | ConvertFrom-Json
    foreach ($k in @('output', 'language_id', 'bookmark_prefix', 'style_prefix', 'batches')) {
        if ($null -eq $c.$k) { throw "Project config is missing '$k'." }
    }
    if ($c.bookmark_prefix -notmatch '^[A-Za-z][A-Za-z0-9]{0,15}$') { throw 'bookmark_prefix must be letters/digits, max 16 characters.' }
    $validation = if ($c.validation) { $c.validation } else { [IO.Path]::GetFileNameWithoutExtension($full) + '-validation.json' }
    [pscustomobject]@{
        ConfigPath         = $full
        Dir                = $dir
        Config             = $c
        Output             = Resolve-ProjectPath $dir $c.output
        Source             = Resolve-ProjectPath $dir $c.source_file
        SourceSha256       = [string]$c.source_sha256
        AssetsDir          = Resolve-ProjectPath $dir $(if ($c.assets_dir) { $c.assets_dir } else { '.' })
        Batches            = @($c.batches | ForEach-Object { Resolve-ProjectPath $dir $_ })
        Validation         = Resolve-ProjectPath $dir $validation
        LanguageId         = [int]$c.language_id
        OriginalLanguageId = $(if ($c.original_language_id) { [int]$c.original_language_id } else { [int]$c.language_id })
        Prefix             = [string]$c.bookmark_prefix
        StylePrefix        = [string]$c.style_prefix
        Font               = $(if ($c.font) { [string]$c.font } else { 'Cambria' })
        FontSize           = $(if ($c.font_size) { [double]$c.font_size } else { 11 })
        TocTitle           = $(if ($c.toc_title) { [string]$c.toc_title } else { 'Obsah' })
        TocLevels          = $(if ($c.toc_levels) { [string]$c.toc_levels } else { '1-2' })
        PageBreakBefore    = $(if ($null -ne $c.page_break_before) { @($c.page_break_before) } else { @('h1', 'h2') })
        ImageWidthCm       = $(if ($c.image_width_cm) { [double]$c.image_width_cm } else { 12 })
        TitlePage          = $c.title_page
    }
}

function Get-BookBlocks($Project) {
    $blocks = New-Object System.Collections.Generic.List[object]
    foreach ($file in $Project.Batches) {
        if (-not (Test-Path -LiteralPath $file)) { throw "Missing batch: $file" }
        $data = Get-Content -LiteralPath $file -Raw -Encoding utf8 | ConvertFrom-Json
        if ($Project.SourceSha256 -and $data.source_sha256 -and $data.source_sha256 -ne $Project.SourceSha256) { throw "Batch was made from another source: $file" }
        if ($data.status -eq 'todo') { throw "Batch is still a skeleton (status todo): $file" }
        foreach ($b in $data.blocks) {
            if ($b.type -notin $script:BlockTypes) { throw "Unsupported block type '$($b.type)' in $file" }
            if ($b.type -eq 'table') {
                if ($b.rows.Count -lt 1 -or $b.rows[0].Count -lt 1) { throw "Empty table in $file" }
                foreach ($row in $b.rows) { if ($row.Count -ne $b.rows[0].Count) { throw "Nonrectangular table in $file" } }
            } elseif ($null -eq $b.text) { throw "Block without text in $file" }
            if ($b.type -eq 'image' -and -not (Test-Path -LiteralPath (Join-Path $Project.AssetsDir $b.file))) { throw "Missing image $($b.file)" }
            $b | Add-Member -NotePropertyName _batch -NotePropertyValue ([IO.Path]::GetFileName($file)) -Force
            $blocks.Add($b)
        }
    }
    return , $blocks.ToArray()
}

function Get-BlockBookmark($Project, [int]$Index) { '{0}_b{1:D4}' -f $Project.Prefix, $Index }

function Test-SourceUnchanged($Project) {
    if ($Project.Source -and $Project.SourceSha256) {
        if ((Get-FileHash -LiteralPath $Project.Source -Algorithm SHA256).Hash -ne $Project.SourceSha256) { throw 'Source file hash mismatch (source missing or changed).' }
        return $true
    }
    return $null
}

function New-OwnedWord {
    $word = New-Object -ComObject Word.Application
    if ($word.Documents.Count -ne 0) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        throw 'The Word instance already contains documents (user session); refusing to automate it.'
    }
    $word.Visible = $false
    $word.DisplayAlerts = 0
    return $word
}

function Close-OwnedWord($Word, $Doc) {
    try { if ($null -ne $Doc) { $Doc.Close(0) } } catch { }
    try { if ($null -ne $Word) { $Word.Quit() } } catch { }
    if ($null -ne $Word) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Word) }
}

function Get-StyleName($Project, [string]$Kind) { $Project.StylePrefix + $Kind }

function Initialize-BookDocument($Doc, $Word, $Project) {
    $ps = $Doc.PageSetup
    $ps.PageWidth = $Word.CentimetersToPoints(21); $ps.PageHeight = $Word.CentimetersToPoints(29.7)
    $ps.TopMargin = $Word.CentimetersToPoints(2.2); $ps.BottomMargin = $Word.CentimetersToPoints(2.2)
    $ps.LeftMargin = $Word.CentimetersToPoints(2.5); $ps.RightMargin = $Word.CentimetersToPoints(2.5)
    $normal = $Doc.Styles.Item(-1)
    # PowerShell 7 caches the argument type of a COM property setter; always pass point values as [double].
    $normal.Font.Name = $Project.Font; $normal.Font.Size = [double]$Project.FontSize
    $normal.LanguageID = $Project.LanguageId
    $normal.ParagraphFormat.SpaceAfter = [double]7; $normal.ParagraphFormat.WidowControl = -1
    foreach ($kind in $script:TextStyles) {
        $s = $Doc.Styles.Add((Get-StyleName $Project $kind), 1)
        $s.BaseStyle = $normal.NameLocal
        $s.LanguageID = $Project.LanguageId
        $s.ParagraphFormat.OutlineLevel = 10
        $s.ParagraphFormat.PageBreakBefore = 0; $s.ParagraphFormat.KeepWithNext = 0; $s.ParagraphFormat.KeepTogether = 0
        $s.ParagraphFormat.Alignment = 3
    }
    $st = { param($k) $Doc.Styles.Item((Get-StyleName $Project $k)) }
    (& $st 'Note').Font.Size = [double]9; (& $st 'Note').ParagraphFormat.Alignment = 0
    (& $st 'Label').Font.Bold = -1; (& $st 'Label').ParagraphFormat.KeepWithNext = -1; (& $st 'Label').ParagraphFormat.Alignment = 0
    (& $st 'Quote').Font.Italic = -1; (& $st 'Quote').ParagraphFormat.Alignment = 1
    (& $st 'Caption').ParagraphFormat.Alignment = 1; (& $st 'Caption').ParagraphFormat.KeepTogether = -1
    (& $st 'Signature').ParagraphFormat.Alignment = 2
    foreach ($k in @('Author', 'OriginalTitle')) { (& $st $k).ParagraphFormat.Alignment = 1; (& $st $k).LanguageID = $Project.OriginalLanguageId }
    (& $st 'Author').Font.Size = [double]13; (& $st 'Author').ParagraphFormat.SpaceBefore = [double]18
    (& $st 'OriginalTitle').Font.Size = [double]15; (& $st 'OriginalTitle').Font.Bold = -1
    (& $st 'Edition').ParagraphFormat.Alignment = 1; (& $st 'Edition').LanguageID = $Project.OriginalLanguageId
    $sizes = @{ -2 = 18; -3 = 15; -4 = 12 }
    foreach ($id in @(-2, -3, -4)) {
        $s = $Doc.Styles.Item($id)
        $s.Font.Name = $Project.Font; $s.Font.Size = [double]$sizes[$id]; $s.Font.Bold = -1
        $s.LanguageID = $Project.LanguageId
        $s.ParagraphFormat.KeepWithNext = -1
        $s.ParagraphFormat.SpaceBefore = [double]$(if ($id -eq -4) { 12 } else { 0 })
        $s.ParagraphFormat.SpaceAfter = [double]$(if ($id -eq -4) { 6 } else { 12 })
        $level = @{ -2 = 'h1'; -3 = 'h2'; -4 = 'h3' }[$id]
        $s.ParagraphFormat.PageBreakBefore = $(if ($Project.PageBreakBefore -contains $level) { -1 } else { 0 })
    }
    foreach ($id in @(-20, -21, -22)) {
        $s = $Doc.Styles.Item($id)
        $s.Font.Name = $Project.Font; $s.Font.Size = [double]$Project.FontSize
        $s.ParagraphFormat.SpaceAfter = [double]5;
    }
    foreach ($id in @(-63, -75)) { $Doc.Styles.Item($id).Font.Name = $Project.Font; $Doc.Styles.Item($id).ParagraphFormat.Alignment = 1 }
    $Doc.Styles.Item(-63).Font.Size = [double]26
}

function Add-BookParagraph($Doc, [string]$Text, $Style, [int]$Language, [string]$Bookmark = '') {
    $start = $Doc.Content.End - 1
    $Text = $Text.Replace("`r`n", "`n").Replace("`n", "`v")
    $r = $Doc.Range($start, $start)
    $r.Font.Reset(); $r.ParagraphFormat.Reset()
    $r.Text = $Text + "`r"
    $r = $Doc.Range($start, $start + $Text.Length + 1)
    $r.Style = $Style
    $r.LanguageID = $Language
    if ($Bookmark) { $Doc.Bookmarks.Add($Bookmark, $Doc.Range($start, $start + $Text.Length)) | Out-Null }
    return $r
}

function Get-BlockLanguage($Project, $Block) { if ($Block.lang) { [int]$Block.lang } else { $Project.LanguageId } }

function Add-BookBlock($Doc, $Word, $Project, $Block, [string]$Bookmark) {
    $lang = Get-BlockLanguage $Project $Block
    if ($Block.type -eq 'table') {
        $start = $Doc.Content.End - 1
        $rows = $Block.rows.Count; $cols = $Block.rows[0].Count
        $table = $Doc.Tables.Add($Doc.Range($start, $start), $rows, $cols)
        $table.Range.Font.Reset(); $table.Range.ParagraphFormat.Reset()
        $table.Range.Style = Get-StyleName $Project 'Text'
        $table.Range.Font.Size = [double][Math]::Max(8, $Project.FontSize - 1)
        $table.Range.ParagraphFormat.Alignment = 0; $table.Range.ParagraphFormat.SpaceAfter = [double]3
        $table.Range.ParagraphFormat.PageBreakBefore = 0; $table.Range.ParagraphFormat.KeepWithNext = 0
        $table.Rows.AllowBreakAcrossPages = 0
        $table.Rows.Item(1).HeadingFormat = -1
        $table.Rows.Item(1).Range.Font.Bold = -1
        $table.Borders.Enable = 1
        $table.AutoFitBehavior(0)
        $widths = if ($Block.col_widths_cm -and $Block.col_widths_cm.Count -eq $cols) { @($Block.col_widths_cm) } else { @(1..$cols | ForEach-Object { 16 / $cols }) }
        for ($c = 1; $c -le $cols; $c++) { $table.Columns.Item($c).Width = $Word.CentimetersToPoints([double]$widths[$c - 1]) }
        for ($row = 1; $row -le $rows; $row++) {
            for ($col = 1; $col -le $cols; $col++) { $table.Cell($row, $col).Range.Text = [string]$Block.rows[$row - 1][$col - 1] }
        }
        $table.Range.LanguageID = $lang
        $Doc.Bookmarks.Add($Bookmark, $table.Range) | Out-Null
        $Doc.Range($table.Range.End, $table.Range.End).InsertParagraphAfter()
        return
    }
    $caption = Get-StyleName $Project 'Caption'
    if ($Block.type -eq 'image') {
        $start = $Doc.Content.End - 1
        $r = $Doc.Range($start, $start)
        $r.Font.Reset(); $r.ParagraphFormat.Reset()
        $r.Text = "`r"
        $r = $Doc.Range($start, $start + 1)
        $r.Style = $caption
        $r.ParagraphFormat.PageBreakBefore = $(if ($Block.new_page) { -1 } else { 0 })
        $r.ParagraphFormat.KeepWithNext = -1
        $shape = $Doc.InlineShapes.AddPicture((Join-Path $Project.AssetsDir $Block.file), $false, $true, $Doc.Range($start, $start))
        $shape.LockAspectRatio = -1
        $shape.Width = $Word.CentimetersToPoints($(if ($Block.width_cm) { [double]$Block.width_cm } else { $Project.ImageWidthCm }))
        $shape.AlternativeText = [string]$Block.text
        Add-BookParagraph $Doc ([string]$Block.text) $caption $lang $Bookmark | Out-Null
        return
    }
    $style = switch ($Block.type) {
        'h1' { $Doc.Styles.Item(-2).NameLocal }
        'h2' { $Doc.Styles.Item(-3).NameLocal }
        'h3' { $Doc.Styles.Item(-4).NameLocal }
        'p' { Get-StyleName $Project 'Text' }
        'note' { Get-StyleName $Project 'Note' }
        'label' { Get-StyleName $Project 'Label' }
        'quote' { Get-StyleName $Project 'Quote' }
        'signature' { Get-StyleName $Project 'Signature' }
    }
    $r = Add-BookParagraph $Doc ([string]$Block.text) $style $lang $Bookmark
    if ($null -ne $Block.new_page) { $r.ParagraphFormat.PageBreakBefore = $(if ($Block.new_page) { -1 } else { 0 }) }
    if ($Block.type -eq 'p' -and ([string]$Block.text).Contains("`n")) { $r.ParagraphFormat.Alignment = 0 }
}

# Title pages required by the Translator agent (layout approved by the user, all centred):
# page 1: title (Title style) + subtitle (Subtitle style) in the document language, horizontal rule, original title
# (Title) + subtitle (Subtitle) when this is a translation, blank line, author(s) (Subtitle), author_details (Text);
# page 2: edition lines from the original ('' = blank line), blank line, AI note (translations only), blank line, notes.
function Add-BookTitlePages($Doc, $Project) {
    $t = $Project.TitlePage
    if (-not $t) { throw 'Project config is missing title_page.' }
    if (-not $t.title -or -not $t.authors -or -not $t.edition) { throw 'title_page needs title, authors and edition.' }
    $isTranslation = [bool]$t.original_title
    if ($isTranslation -and -not $t.ai_note) { throw 'A translation needs title_page.ai_note ("Automatický překlad pomocí <model> z originálu.").' }
    if ($t.ai_note -and $t.ai_note -match '<|>') { throw 'Replace the AI model placeholder in title_page.ai_note with the real model name.' }
    if ($t.ai_note -and $Project.LanguageId -eq 1029 -and $t.ai_note -notmatch '^Automatický překlad pomocí \S.* z originálu\.$') { throw 'title_page.ai_note must read "Automatický překlad pomocí <model> z originálu."' }
    $p = $Project.Prefix; $lang = $Project.LanguageId; $orig = $Project.OriginalLanguageId
    Add-BookParagraph $Doc ([string]$t.title) $Doc.Styles.Item(-63).NameLocal $(if ($isTranslation) { $lang } else { $orig }) "${p}_title" | Out-Null
    if ($t.subtitle) { Add-BookParagraph $Doc ([string]$t.subtitle) $Doc.Styles.Item(-75).NameLocal $(if ($isTranslation) { $lang } else { $orig }) "${p}_subtitle" | Out-Null }
    $r = Add-BookParagraph $Doc '' (Get-StyleName $Project 'Author') $orig
    $rule = $r.ParagraphFormat.Borders.Item(-3)
    $rule.LineStyle = 1; $rule.LineWidth = 12; $rule.Color = 0
    $r.ParagraphFormat.SpaceBefore = [double]6; $r.ParagraphFormat.SpaceAfter = [double]12
    $Doc.Bookmarks.Add("${p}_rule", $r) | Out-Null
    if ($isTranslation) {
        Add-BookParagraph $Doc ([string]$t.original_title) $Doc.Styles.Item(-63).NameLocal $orig "${p}_original_title" | Out-Null
        if ($t.original_subtitle) { Add-BookParagraph $Doc ([string]$t.original_subtitle) $Doc.Styles.Item(-75).NameLocal $orig "${p}_original_subtitle" | Out-Null }
    }
    $text = Get-StyleName $Project 'Text'
    $center = { param($r) $r.ParagraphFormat.Alignment = 1; $r }
    & $center (Add-BookParagraph $Doc '' $text $orig) | Out-Null
    $i = 0
    foreach ($a in @($t.authors)) { $i++; Add-BookParagraph $Doc ([string]$a) $Doc.Styles.Item(-75).NameLocal $orig ('{0}_author{1:D2}' -f $p, $i) | Out-Null }
    $i = 0
    foreach ($a in @($t.author_details)) { if (-not $a) { continue }; $i++; & $center (Add-BookParagraph $Doc ([string]$a) $text $orig ('{0}_author_details{1:D2}' -f $p, $i)) | Out-Null }
    $i = 0
    foreach ($e in @($t.edition)) {
        $i++
        $r = Add-BookParagraph $Doc ([string]$e) (Get-StyleName $Project 'Edition') $orig $(if ($e) { '{0}_edition{1:D2}' -f $p, $i } else { '' })
        if ($i -eq 1) { $r.ParagraphFormat.PageBreakBefore = -1 }
    }
    if ($t.ai_note) {
        & $center (Add-BookParagraph $Doc '' $text $lang) | Out-Null
        & $center (Add-BookParagraph $Doc ([string]$t.ai_note) $text $lang "${p}_ai_note") | Out-Null
    }
    if (@($t.notes | Where-Object { $_ }).Count) { Add-BookParagraph $Doc '' (Get-StyleName $Project 'Note') $lang | Out-Null }
    $i = 0
    foreach ($n in @($t.notes)) {
        if (-not $n) { continue }
        $i++
        $text = if ($n -is [string]) { $n } else { [string]$n.text }
        $nl = if ($n -isnot [string] -and $n.lang) { [int]$n.lang } else { $lang }
        $bm = if ($n -isnot [string] -and $n.id) { "${p}_$($n.id)" } else { '{0}_front{1:D2}' -f $p, $i }
        Add-BookParagraph $Doc $text (Get-StyleName $Project 'Note') $nl $bm | Out-Null
    }
}

function Add-BookToc($Doc, $Project) {
    $r = Add-BookParagraph $Doc $Project.TocTitle $Doc.Styles.Item(-1).NameLocal $Project.LanguageId "$($Project.Prefix)_toc_heading"
    $r.Font.Size = [double]20; $r.Font.Bold = -1
    $r.ParagraphFormat.PageBreakBefore = -1; $r.ParagraphFormat.OutlineLevel = 10
    $start = $Doc.Content.End - 1
    Add-BookParagraph $Doc '' $Doc.Styles.Item(-1).NameLocal $Project.LanguageId | Out-Null
    $lv = $Project.TocLevels -split '-'
    $toc = $Doc.TablesOfContents.Add($Doc.Range($start, $start), $true, [int]$lv[0], [int]$lv[-1])
    $toc.UseHyperlinks = $true; $toc.IncludePageNumbers = $true; $toc.RightAlignPageNumbers = $true
    return $toc
}

function Set-BookFooter($Doc) {
    $footer = $Doc.Sections.Item(1).Footers.Item(1).Range
    $footer.Fields.Add($footer, 33) | Out-Null
    $footer = $Doc.Sections.Item(1).Footers.Item(1).Range
    $footer.Paragraphs.Item(1).Alignment = 1
}

function Update-BookToc($Doc, $Project) {
    $toc = $Doc.TablesOfContents.Item(1)
    foreach ($field in $Doc.Fields) {
        if ($field.Type -eq 13) { $field.Code.Text = " TOC \o `"$($Project.TocLevels)`" \h \z " }
    }
    $last = ''
    for ($i = 0; $i -lt 10; $i++) {
        $Doc.Repaginate(); $toc.Update(); $Doc.Repaginate(); $toc.UpdatePageNumbers()
        if ($toc.Range.Text -ceq $last) { return $toc }
        $last = $toc.Range.Text
    }
    throw 'TOC pagination did not stabilize.'
}

# Verifies an opened document against the block list. Returns a report object.
# User revisions are reported in text_differences (never silently "fixed").
function Test-BookDocument($Doc, $Project, $Blocks, [switch]$SkipTitlePages) {
    $p = $Project.Prefix
    $problems = New-Object System.Collections.Generic.List[string]
    $differences = New-Object System.Collections.Generic.List[object]
    $langWarnings = 0
    $langBlocks = New-Object System.Collections.Generic.List[object]
    if (-not $SkipTitlePages -and $Project.TitlePage) {
        foreach ($b in @("${p}_title", "${p}_rule", "${p}_author01", "${p}_edition01")) { if (-not $Doc.Bookmarks.Exists($b)) { $problems.Add("Missing title-page element $b") } }
        if ($Project.TitlePage.original_title) { foreach ($b in @("${p}_original_title", "${p}_ai_note")) { if (-not $Doc.Bookmarks.Exists($b)) { $problems.Add("Missing title-page element $b") } } }
        if ($Doc.Bookmarks.Exists("${p}_title") -and $Doc.Bookmarks.Item("${p}_title").Range.Information(3) -ne 1) { $problems.Add('Title is not on page 1.') }
        if ($Doc.Bookmarks.Exists("${p}_rule") -and $Doc.Bookmarks.Item("${p}_rule").Range.ParagraphFormat.Borders.Item(-3).LineStyle -eq 0) { $problems.Add('Horizontal rule missing.') }
        if ($Doc.Bookmarks.Exists("${p}_edition01")) {
            $ep = $Doc.Bookmarks.Item("${p}_edition01").Range.Information(3)
            if ($Doc.Bookmarks.Exists("${p}_author01") -and $ep -le $Doc.Bookmarks.Item("${p}_author01").Range.Information(3)) { $problems.Add('Edition details are not on a new page.') }
            if ($Doc.Bookmarks.Exists("${p}_ai_note") -and $Doc.Bookmarks.Item("${p}_ai_note").Range.Information(3) -ne $ep) { $problems.Add('AI note is not on the edition page.') }
        }
    }
    $cells = 0; $images = 0
    for ($i = 1; $i -le $Blocks.Count; $i++) {
        $b = $Blocks[$i - 1]; $name = Get-BlockBookmark $Project $i
        if (-not $Doc.Bookmarks.Exists($name)) { $problems.Add("Missing block bookmark $name"); continue }
        $range = $Doc.Bookmarks.Item($name).Range
        if ($b.type -eq 'table') {
            if ($range.Tables.Count -lt 1) { $problems.Add("Table missing at $name"); continue }
            $table = $range.Tables.Item(1)
            if ($table.Rows.Count -ne $b.rows.Count -or $table.Columns.Count -ne $b.rows[0].Count) { $problems.Add("Table dimensions differ at $name"); continue }
            for ($r = 1; $r -le $b.rows.Count; $r++) {
                for ($c = 1; $c -le $b.rows[0].Count; $c++) {
                    $actual = $table.Cell($r, $c).Range.Text.TrimEnd([char[]]@([char]13, [char]7))
                    $cells++
                    if ($actual -cne [string]$b.rows[$r - 1][$c - 1]) { $differences.Add([pscustomobject]@{ block = $name; cell = "$r/$c"; expected = [string]$b.rows[$r - 1][$c - 1]; actual = $actual }) }
                }
            }
            continue
        }
        $expected = ([string]$b.text).Replace("`r`n", "`n").Replace("`n", "`v")
        if ($range.Text -cne $expected) { $differences.Add([pscustomobject]@{ block = $name; type = $b.type; expected = $expected; actual = $range.Text }) }
        if ($range.LanguageID -ne (Get-BlockLanguage $Project $b)) { $langWarnings++; $langBlocks.Add([pscustomobject]@{ block = $name; language_id = $range.LanguageID; text = ([string]$b.text).Substring(0, [Math]::Min(80, ([string]$b.text).Length)) }) }
        if ($b.type -eq 'image') {
            $images++
            $prev = $range.Paragraphs.Item(1).Range.Previous(4, 1)
            for ($k = 0; $k -lt 2 -and $null -ne $prev -and $prev.InlineShapes.Count -lt 1 -and -not $prev.Text.Trim(); $k++) { $prev = $prev.Previous(4, 1) }
            if ($null -eq $prev -or $prev.InlineShapes.Count -lt 1) { $problems.Add("Image missing before caption $name") }
            elseif ($prev.InlineShapes.Item(1).Range.Information(3) -ne $range.Information(3)) { $problems.Add("Image and caption split across pages at $name") }
        }
    }
    $toc = $Doc.TablesOfContents.Item(1)
    $maxLevel = [int](($Project.TocLevels -split '-')[-1])
    $expectedHeadings = @(for ($i = 1; $i -le $Blocks.Count; $i++) { if ($Blocks[$i - 1].type -match '^h(\d)$' -and [int]$Matches[1] -le $maxLevel) { Get-BlockBookmark $Project $i } })
    $links = New-Object System.Collections.Generic.List[object]
    foreach ($link in $toc.Range.Hyperlinks) {
        $target = [string]$link.SubAddress
        if (-not $target -or -not $Doc.Bookmarks.Exists($target)) { $problems.Add("Broken TOC link: $target"); continue }
        $parts = $link.Range.Text.Trim() -split "`t"
        $title = $parts[0].Trim()
        $shown = if ($parts.Count -ge 2) { [int]$parts[-1] } else { -1 }
        $dest = $Doc.Bookmarks.Item($target).Range.Duplicate; $dest.Collapse(1)
        $page = [int]$dest.Information(3)
        if ($shown -ne $page) { $problems.Add("TOC page mismatch: $title shows $shown, actual $page") }
        $heading = $Doc.Bookmarks.Item($target).Range.Paragraphs.Item(1).Range.Text.Trim()
        if ($heading -ne $title) { $problems.Add("TOC text differs from heading: $title") }
        if ($title -eq $Project.TocTitle) { $problems.Add('TOC lists itself.') }
        $links.Add([pscustomobject]@{ title = $title; page = $page; bookmark = $target })
    }
    if ($links.Count -ne $expectedHeadings.Count) { $problems.Add("TOC has $($links.Count) entries, expected $($expectedHeadings.Count) headings.") }
    else {
        for ($i = 0; $i -lt $links.Count; $i++) {
            $h = $Doc.Bookmarks.Item($expectedHeadings[$i]).Range.Paragraphs.Item(1).Range.Text.Trim()
            if ($links[$i].title -ne $h) { $problems.Add("TOC order differs at entry $($i + 1): $($links[$i].title)"); break }
        }
    }
    [pscustomobject]@{
        blocks_checked = $Blocks.Count
        table_cells_checked = $cells
        images_checked = $images
        inline_shapes = $Doc.InlineShapes.Count
        text_differences = $differences.ToArray()
        language_warnings = $langWarnings
        language_warning_blocks = $langBlocks.ToArray()
        toc_links = $links.ToArray()
        document_pages = $Doc.ComputeStatistics(2)
        revisions = $Doc.Revisions.Count
        comments = $Doc.Comments.Count
        track_revisions = [bool]$Doc.TrackRevisions
        problems = $problems.ToArray()
    }
}

# ZIP/XML integrity; optionally confirms that every media part of a baseline package survived unchanged.
function Test-DocxPackage([string]$Path, [string]$Baseline = '') {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    $base = $null
    try {
        $xmlParts = 0
        foreach ($e in $zip.Entries | Where-Object { $_.FullName -match '\.(xml|rels)$' }) {
            $s = $e.Open(); try { (New-Object System.Xml.XmlDocument).Load($s); $xmlParts++ } finally { $s.Dispose() }
        }
        $docCodes = ''; $footerCodes = ''
        foreach ($e in $zip.Entries | Where-Object { $_.FullName -match '^word/(document|footer\d*)\.xml$' }) {
            $s = $e.Open()
            try { $x = New-Object System.Xml.XmlDocument; $x.Load($s) } finally { $s.Dispose() }
            $ns = New-Object System.Xml.XmlNamespaceManager($x.NameTable)
            $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            $codes = (@($x.SelectNodes('//w:instrText', $ns) | ForEach-Object { $_.InnerText }) + @($x.SelectNodes('//w:fldSimple/@w:instr', $ns) | ForEach-Object { $_.Value })) -join ' '
            if ($e.FullName -eq 'word/document.xml') { $docCodes = $codes } else { $footerCodes += ' ' + $codes }
        }
        if ($docCodes -notmatch '\bTOC\b' -or $docCodes -notmatch '\\h') { throw 'Hyperlinked TOC field (\h) not present.' }
        if ($footerCodes -notmatch '\bPAGE\b') { throw 'PAGE field missing in footer.' }
        $media = @($zip.Entries | Where-Object { $_.FullName -like 'word/media/*' })
        if ($Baseline) {
            $base = [IO.Compression.ZipFile]::OpenRead($Baseline)
            $sha = [Security.Cryptography.SHA256]::Create()
            try {
                foreach ($e in $base.Entries | Where-Object { $_.FullName -like 'word/media/*' }) {
                    $o = $zip.GetEntry($e.FullName)
                    if ($null -eq $o) { throw "Existing media lost: $($e.FullName)" }
                    $a = $e.Open(); $b = $o.Open()
                    try { if ([Convert]::ToHexString($sha.ComputeHash($a)) -ne [Convert]::ToHexString($sha.ComputeHash($b))) { throw "Existing media changed: $($e.FullName)" } } finally { $a.Dispose(); $b.Dispose() }
                }
            } finally { $sha.Dispose() }
        }
        [pscustomobject]@{ xml_parts = $xmlParts; media_parts = $media.Count; toc_hyperlinked = $true }
    } finally { $zip.Dispose(); if ($base) { $base.Dispose() } }
}

function Write-BookReport($Project, $Report) {
    $Report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Project.Validation -Encoding utf8
}

Export-ModuleMember -Function Read-BookProject, Get-BookBlocks, Get-BlockBookmark, Test-SourceUnchanged, New-OwnedWord, Close-OwnedWord,
    Get-StyleName, Initialize-BookDocument, Add-BookParagraph, Add-BookBlock, Add-BookTitlePages, Add-BookToc, Set-BookFooter,
    Update-BookToc, Test-BookDocument, Test-DocxPackage, Write-BookReport
