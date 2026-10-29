<#
.SYNOPSIS
  Local OCR of rendered page images with the built-in Windows OCR engine (Windows.Media.Ocr). No network.
.DESCRIPTION
  Input: PNG files named page-N.png (output of render.mjs). Output: ocr-N.json per page with lines,
  words and pixel bounding boxes, and ocr-N.txt with plain lines. Existing outputs are kept unless -Overwrite.
  WinRT projections need Windows PowerShell 5.1; under PowerShell 7 the script re-launches itself in powershell.exe.
  Recognised text is still only a draft: verify it against the page image.
.EXAMPLE
  node render.mjs book.pdf renders 30-34 --scale 2.5
  .\ocr-windows.ps1 -ImagesDir renders -OutDir ocr -Pages 30-34 -Language en-US
#>
param(
  [Parameter(Mandatory)] [string] $ImagesDir,
  [Parameter(Mandatory)] [string] $OutDir,
  [string] $Pages,
  [string] $Language = 'en-US',
  [switch] $Overwrite,
  [switch] $ListLanguages
)
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSEdition -eq 'Core') {
  $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-ImagesDir', (Resolve-Path $ImagesDir).Path, '-OutDir', $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir), '-Language', $Language)
  if ($Pages) { $argList += @('-Pages', $Pages) }
  if ($Overwrite) { $argList += '-Overwrite' }
  if ($ListLanguages) { $argList += '-ListLanguages' }
  & powershell.exe @argList
  exit $LASTEXITCODE
}

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime]
$asTaskGeneric = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
  $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
function Await($op, [type] $type) {
  $task = $asTaskGeneric.MakeGenericMethod($type).Invoke($null, @($op))
  $null = $task.Wait(-1)
  $task.Result
}

if ($ListLanguages) { [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages | ForEach-Object { $_.LanguageTag }; exit 0 }
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new($Language))
if (-not $engine) { throw "Windows OCR language '$Language' is not installed. Available: $(([Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages | ForEach-Object LanguageTag) -join ', ')" }

$ImagesDir = (Resolve-Path $ImagesDir).Path
New-Item -ItemType Directory -Force $OutDir | Out-Null
$files = Get-ChildItem $ImagesDir -Filter 'page-*.png' | Where-Object { $_.BaseName -match '^page-(\d+)$' } |
  ForEach-Object { [pscustomobject]@{ N = [int]$Matches[1]; File = $_ } }
if ($Pages) {
  $wanted = foreach ($part in $Pages -split ',') { $a, $b = $part.Trim() -split '-'; if ($b) { [int]$a..[int]$b } else { [int]$a } }
  $missing = $wanted | Where-Object { $_ -notin $files.N }
  if ($missing) { throw "Missing rendered pages: $($missing -join ', ')" }
  $files = $files | Where-Object { $_.N -in $wanted }
}
foreach ($f in ($files | Sort-Object N)) {
  $json = Join-Path $OutDir "ocr-$($f.N).json"
  if ((Test-Path $json) -and -not $Overwrite) { Write-Host "$($f.N): exists, kept"; continue }
  $sf = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($f.File.FullName)) ([Windows.Storage.StorageFile])
  $stream = Await ($sf.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
  try {
    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    if ([Math]::Max($bitmap.PixelWidth, $bitmap.PixelHeight) -gt [Windows.Media.Ocr.OcrEngine]::MaxImageDimension) {
      throw "Page $($f.N): image larger than OcrEngine.MaxImageDimension ($([Windows.Media.Ocr.OcrEngine]::MaxImageDimension)); render with a smaller --scale."
    }
    $result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
  } finally { $stream.Dispose() }
  $lines = foreach ($l in $result.Lines) {
    $words = foreach ($w in $l.Words) { $r = $w.BoundingRect; [pscustomobject][ordered]@{ t = $w.Text; x = [int]$r.X; y = [int]$r.Y; w = [int]$r.Width; h = [int]$r.Height } }
    $x = ($words | Measure-Object -Property x -Minimum).Minimum
    $y = ($words | Measure-Object -Property y -Minimum).Minimum
    $right = ($words | ForEach-Object { $_.x + $_.w } | Measure-Object -Maximum).Maximum
    $bottom = ($words | ForEach-Object { $_.y + $_.h } | Measure-Object -Maximum).Maximum
    [ordered]@{ text = $l.Text; x = $x; y = $y; w = $right - $x; h = $bottom - $y; words = @($words) }
  }
  $page = [ordered]@{ page = $f.N; image = $f.File.Name; width = $bitmap.PixelWidth; height = $bitmap.PixelHeight; angle = $result.TextAngle; language = $Language; lines = @($lines) }
  [IO.File]::WriteAllText($json, ($page | ConvertTo-Json -Depth 6 -Compress), [Text.UTF8Encoding]::new($false))
  [IO.File]::WriteAllLines((Join-Path $OutDir "ocr-$($f.N).txt"), [string[]]@($lines | ForEach-Object { $_.text }), [Text.UTF8Encoding]::new($false))
  Write-Host "$($f.N): $(@($lines).Count) lines"
}
