param(
  [Parameter(Mandatory = $true)]
  [string]$SourcePath,

  [Parameter(Mandatory = $true)]
  [string]$OutputPath
)

$src = Resolve-Path $SourcePath
$out = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$tmp = Join-Path (Split-Path $src) ('docx_build_' + [guid]::NewGuid().ToString('N'))

New-Item -ItemType Directory -Path $tmp | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp '_rels') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp 'word') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp 'word\_rels') | Out-Null

function Escape-Xml([string]$text) {
  [System.Security.SecurityElement]::Escape($text)
}

function New-Paragraph([string]$text, [string]$style = $null) {
  $styleXml = if ($style) { '<w:pPr><w:pStyle w:val="' + $style + '"/></w:pPr>' } else { '' }
  '<w:p>' + $styleXml + '<w:r><w:t xml:space="preserve">' + (Escape-Xml $text) + '</w:t></w:r></w:p>'
}

function New-Bullet([string]$text) {
  '<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:t xml:space="preserve">' + (Escape-Xml $text) + '</w:t></w:r></w:p>'
}

function New-Numbered([string]$text) {
  '<w:p><w:pPr><w:pStyle w:val="ListNumber"/></w:pPr><w:r><w:t xml:space="preserve">' + (Escape-Xml $text) + '</w:t></w:r></w:p>'
}

function New-Table($rows) {
  $xml = '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/></w:tblBorders></w:tblPr>'
  foreach ($row in $rows) {
    $xml += '<w:tr>'
    foreach ($cell in $row) {
      $xml += '<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>' + (New-Paragraph $cell) + '</w:tc>'
    }
    $xml += '</w:tr>'
  }
  $xml + '</w:tbl>'
}

$lines = Get-Content -Path $src -Encoding UTF8
$body = New-Object System.Collections.Generic.List[string]
$i = 0

while ($i -lt $lines.Count) {
  $line = $lines[$i]
  if ([string]::IsNullOrWhiteSpace($line)) {
    $i++
    continue
  }

  if ($line.TrimStart().StartsWith('|')) {
    $rows = New-Object System.Collections.Generic.List[object]
    while ($i -lt $lines.Count -and $lines[$i].TrimStart().StartsWith('|')) {
      $cells = $lines[$i].Trim().Trim('|').Split('|') | ForEach-Object { $_.Trim() }
      $isSeparator = $true
      foreach ($cell in $cells) {
        if ($cell -notmatch '^:?-{3,}:?$') {
          $isSeparator = $false
        }
      }
      if (-not $isSeparator) {
        $rows.Add($cells) | Out-Null
      }
      $i++
    }
    if ($rows.Count -gt 0) {
      $body.Add((New-Table $rows)) | Out-Null
    }
    continue
  }

  if ($line -match '^(#{1,6})\s+(.+)$') {
    $level = [Math]::Min($matches[1].Length, 3)
    $body.Add((New-Paragraph $matches[2] ('Heading' + $level))) | Out-Null
  } elseif ($line -match '^\s*-\s+(.+)$') {
    $body.Add((New-Bullet $matches[1])) | Out-Null
  } elseif ($line -match '^\s*\d+\.\s+(.+)$') {
    $body.Add((New-Numbered $matches[1])) | Out-Null
  } else {
    $body.Add((New-Paragraph $line)) | Out-Null
  }
  $i++
}

$contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
$rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
$docRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
$styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="360" w:after="160"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="280" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="220" w:after="100"/></w:pPr><w:rPr><w:b/><w:sz w:val="23"/><w:szCs w:val="23"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style></w:styles>'
$document = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + ($body -join '') + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>'

[System.IO.File]::WriteAllText((Join-Path $tmp '[Content_Types].xml'), $contentTypes, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText((Join-Path $tmp '_rels\.rels'), $rels, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText((Join-Path $tmp 'word\_rels\document.xml.rels'), $docRels, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText((Join-Path $tmp 'word\styles.xml'), $styles, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText((Join-Path $tmp 'word\document.xml'), $document, [System.Text.Encoding]::UTF8)

if (Test-Path $out) {
  Remove-Item -LiteralPath $out -Force
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$zip = [System.IO.Compression.ZipFile]::Open($out, [System.IO.Compression.ZipArchiveMode]::Create)
try {
  [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, (Join-Path $tmp '[Content_Types].xml'), '[Content_Types].xml') | Out-Null
  [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, (Join-Path $tmp '_rels\.rels'), '_rels/.rels') | Out-Null
  [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, (Join-Path $tmp 'word\document.xml'), 'word/document.xml') | Out-Null
  [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, (Join-Path $tmp 'word\styles.xml'), 'word/styles.xml') | Out-Null
  [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, (Join-Path $tmp 'word\_rels\document.xml.rels'), 'word/_rels/document.xml.rels') | Out-Null
} finally {
  $zip.Dispose()
}

Remove-Item -LiteralPath $tmp -Recurse -Force
Get-Item -LiteralPath $out | Select-Object FullName, Length, LastWriteTime
