# Exports a book DOCX to PDF (headings as PDF bookmarks) from a temporary copy, for visual checks of the layout.
# Optional: renders chosen pages to PNG with pdf-source-tools\render.mjs so they can be viewed.
# Usage: & export-pdf.ps1 -Docx book.docx -Pdf book.pdf [-RenderPages "1-3,10"] [-RenderDir pages]
param([Parameter(Mandatory)][string]$Docx, [Parameter(Mandatory)][string]$Pdf, [string]$RenderPages = '', [string]$RenderDir = '')
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'BookDocx.psm1') -Force
$src = (Resolve-Path -LiteralPath $Docx).Path
$pdfPath = [IO.Path]::GetFullPath($Pdf)
if (Test-Path -LiteralPath $pdfPath) { throw "PDF exists; choose another name: $pdfPath" }
$temp = Join-Path ([IO.Path]::GetTempPath()) ("book-pdf-" + [guid]::NewGuid().ToString('N') + '.docx')
Copy-Item -LiteralPath $src -Destination $temp
$word = New-OwnedWord
$doc = $null
try {
    $doc = $word.Documents.Open($temp, $false, $true, $false)
    $doc.Repaginate()
    $pages = $doc.ComputeStatistics(2)
    $doc.ExportAsFixedFormat($pdfPath, 17, $false, 0, 0, 1, 1, 0, $true, $true, 1)
    $doc.Close(0); $doc = $null
} finally { Close-OwnedWord $word $doc; Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
Write-Output "PDF: $pdfPath ($pages pages)"
if ($RenderPages) {
    $dir = if ($RenderDir) { $RenderDir } else { [IO.Path]::ChangeExtension($pdfPath, $null).TrimEnd('.') + '-pages' }
    node (Join-Path $PSScriptRoot '..\pdf-source-tools\render.mjs') $pdfPath $dir $RenderPages
    if ($LASTEXITCODE) { throw 'Rendering failed.' }
}
