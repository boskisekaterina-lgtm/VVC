$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$docxPath = Join-Path $root 'vvs_suppliers_A3_print_layout.docx'
$temp = Join-Path $root '.docx_build'

if (Test-Path -LiteralPath $temp) {
    Remove-Item -LiteralPath $temp -Recurse -Force
}

New-Item -ItemType Directory -Force -Path (Join-Path $temp '_rels') | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $temp 'word') | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $temp 'word\_rels') | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $temp 'word\media') | Out-Null

$images = Get-ChildItem -LiteralPath $root -Filter '*.png' | Sort-Object Name
foreach ($image in $images) {
    Copy-Item -LiteralPath $image.FullName -Destination (Join-Path $temp ('word\media\' + $image.Name)) -Force
}

$contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
'@

$rootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'@

$docRelsItems = @()
$bodyItems = @()
$cx = 15120000
$cy = 10692000
$docPrId = 1

foreach ($image in $images) {
    $rid = 'rId' + $docPrId
    $docRelsItems += "  <Relationship Id=`"$rid`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`" Target=`"media/$($image.Name)`"/>"
    $bodyItems += @"
<w:p>
  <w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>
  <w:r>
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="$cx" cy="$cy"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:docPr id="$docPrId" name="A3 page $docPrId"/>
        <wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:nvPicPr>
                <pic:cNvPr id="$docPrId" name="$($image.Name)"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="$rid" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="$cx" cy="$cy"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>
"@
    if ($docPrId -lt $images.Count) {
        $bodyItems += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    }
    $docPrId++
}

$docRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
$($docRelsItems -join "`n")
</Relationships>
"@

$document = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
$($bodyItems -join "`n")
    <w:sectPr>
      <w:pgSz w:w="23811" w:h="16838" w:orient="landscape"/>
      <w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"@

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText((Join-Path $temp '[Content_Types].xml'), $contentTypes, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $temp '_rels\.rels'), $rootRels, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $temp 'word\_rels\document.xml.rels'), $docRels, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $temp 'word\document.xml'), $document, $utf8NoBom)

if (Test-Path -LiteralPath $docxPath) {
    Remove-Item -LiteralPath $docxPath -Force
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($temp, $docxPath)
Remove-Item -LiteralPath $temp -Recurse -Force
