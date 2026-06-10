# Convert markdown to DOCX via HTML + Word COM
$mdPath   = "c:\Users\ПК\test\ВВС_концепция_обновленная.md"
$htmlPath = "c:\Users\ПК\test\_temp_vvs.html"
$docPath  = "c:\Users\ПК\test\ВВС_концепция_обновленная.docx"

$mdContent = [System.IO.File]::ReadAllText($mdPath, [System.Text.Encoding]::UTF8)
$lines     = $mdContent -split "`r?`n"

# ── HTML builder ──────────────────────────────────────────────────────────────
$parts = [System.Collections.Generic.List[string]]::new()

$parts.Add('<!DOCTYPE html><html><head><meta charset="UTF-8"><style>')
$parts.Add('body{font-family:Calibri,Arial,sans-serif;font-size:11pt;margin:0;}')
$parts.Add('h1{font-size:18pt;color:#1F3864;border-bottom:2px solid #1F3864;padding-bottom:4px;margin-top:24pt;margin-bottom:8pt;}')
$parts.Add('h2{font-size:14pt;color:#2E74B5;margin-top:18pt;margin-bottom:6pt;}')
$parts.Add('h3{font-size:12pt;color:#2E74B5;margin-top:14pt;margin-bottom:4pt;}')
$parts.Add('h4{font-size:11pt;color:#2E74B5;margin-top:10pt;margin-bottom:4pt;}')
$parts.Add('h5{font-size:11pt;color:#595959;font-style:italic;margin-top:8pt;margin-bottom:2pt;}')
$parts.Add('h6{font-size:10pt;color:#595959;margin-top:6pt;margin-bottom:2pt;}')
$parts.Add('p{margin:4px 0 8px 0;line-height:1.4;}')
$parts.Add('table{border-collapse:collapse;width:100%;margin:8px 0;font-size:10pt;}')
$parts.Add('td,th{border:1px solid #999;padding:4px 8px;vertical-align:top;}')
$parts.Add('th{background-color:#BDD7EE;font-weight:bold;}')
$parts.Add('pre{font-family:"Courier New";font-size:9pt;background:#F2F2F2;padding:8px;border:1px solid #ddd;white-space:pre-wrap;margin:6px 0;}')
$parts.Add('code{font-family:"Courier New";font-size:9pt;background:#F5F5F5;padding:1px 3px;}')
$parts.Add('blockquote{border-left:4px solid #2E74B5;margin:8px 0;padding:6px 12px;background:#F0F7FF;font-style:italic;font-weight:bold;}')
$parts.Add('blockquote p{margin:0;}')
$parts.Add('hr{border:none;border-top:1px solid #ccc;margin:14px 0;}')
$parts.Add('ul,ol{margin:4px 0 4px 0;padding-left:24px;}')
$parts.Add('li{margin:2px 0;line-height:1.4;}')
$parts.Add('strong{font-weight:bold;}em{font-style:italic;}')
$parts.Add('</style></head><body>')

# ── Inline formatter ──────────────────────────────────────────────────────────
function Escape-Html ([string]$t) {
    $t = $t -replace '&','&amp;'
    $t = $t -replace '<','&lt;'
    $t = $t -replace '>','&gt;'
    return $t
}

function Fmt ([string]$text) {
    $t = Escape-Html $text
    $t = [regex]::Replace($t, '\*\*\*(.+?)\*\*\*', '<strong><em>$1</em></strong>')
    $t = [regex]::Replace($t, '\*\*(.+?)\*\*',     '<strong>$1</strong>')
    $t = [regex]::Replace($t, '(?<!\*)\*([^*\n]+?)\*(?!\*)', '<em>$1</em>')
    $t = [regex]::Replace($t, '`([^`]+?)`',         '<code>$1</code>')
    return $t
}

# ── State ─────────────────────────────────────────────────────────────────────
$inCode    = $false
$codeBuf   = [System.Collections.Generic.List[string]]::new()
$inTable   = $false
$tableBuf  = [System.Collections.Generic.List[string]]::new()
$inList    = $false
$listTag   = ""
$paraBuf   = [System.Text.StringBuilder]::new()

function Flush-Para {
    if ($paraBuf.Length -gt 0) {
        $parts.Add("<p>$($paraBuf.ToString())</p>")
        $paraBuf.Clear() | Out-Null
    }
}

function Flush-List {
    if ($script:inList) {
        $parts.Add("</$($script:listTag)>")
        $script:inList  = $false
        $script:listTag = ""
    }
}

function Flush-Table {
    if ($tableBuf.Count -eq 0) { return }
    $parts.Add('<table>')
    $firstRow = $true
    foreach ($tl in $tableBuf) {
        # separator row: |---|---| or | :--- | :---: |
        if ($tl -match '^\|\s*[-:\s|]+\s*$') { continue }
        $cells = ($tl -split '\|') | Where-Object { $_ -ne $null }
        # trim leading/trailing empty strings from split
        if ($cells.Count -ge 2) { $cells = $cells[1..($cells.Count-2)] }
        if ($firstRow) {
            $parts.Add('<tr>')
            foreach ($c in $cells) { $parts.Add("<th>$(Fmt $c.Trim())</th>") }
            $parts.Add('</tr>')
            $firstRow = $false
        } else {
            $parts.Add('<tr>')
            foreach ($c in $cells) { $parts.Add("<td>$(Fmt $c.Trim())</td>") }
            $parts.Add('</tr>')
        }
    }
    $parts.Add('</table>')
    $tableBuf.Clear()
    $script:inTable = $false
}

# ── Main parse loop ───────────────────────────────────────────────────────────
foreach ($line in $lines) {

    # Code fence
    if ($line -match '^```') {
        if ($inCode) {
            Flush-Para; Flush-List
            $raw = Escape-Html ($codeBuf -join "`n")
            $parts.Add("<pre>$raw</pre>")
            $codeBuf.Clear(); $inCode = $false
        } else {
            Flush-Para; Flush-List; Flush-Table
            $inCode = $true
        }
        continue
    }
    if ($inCode) { $codeBuf.Add($line) | Out-Null; continue }

    # Table row
    if ($line -match '^\|') {
        Flush-Para; Flush-List
        $tableBuf.Add($line) | Out-Null
        $inTable = $true
        continue
    } elseif ($inTable) { Flush-Table }

    # Empty line
    if ($line -match '^\s*$') { Flush-Para; Flush-List; continue }

    # Horizontal rule
    if ($line -match '^-{3,}$') { Flush-Para; Flush-List; $parts.Add('<hr/>'); continue }

    # Headings
    if ($line -match '^(#{1,6})\s+(.+)$') {
        Flush-Para; Flush-List
        $lvl  = $matches[1].Length
        $text = Fmt $matches[2]
        $parts.Add("<h$lvl>$text</h$lvl>")
        continue
    }

    # Blockquote
    if ($line -match '^>\s*(.*)$') {
        Flush-Para; Flush-List
        $parts.Add("<blockquote><p>$(Fmt $matches[1])</p></blockquote>")
        continue
    }

    # Bullet list (- or *)
    if ($line -match '^(\s*)[-*]\s+(.+)$') {
        Flush-Para
        $text = Fmt $matches[2]
        if (-not $inList -or $listTag -ne 'ul') {
            Flush-List
            $parts.Add('<ul>'); $inList = $true; $listTag = 'ul'
        }
        $parts.Add("<li>$text</li>")
        continue
    }

    # Numbered list
    if ($line -match '^\s*\d+\.\s+(.+)$') {
        Flush-Para
        $text = Fmt $matches[1]
        if (-not $inList -or $listTag -ne 'ol') {
            Flush-List
            $parts.Add('<ol>'); $inList = $true; $listTag = 'ol'
        }
        $parts.Add("<li>$text</li>")
        continue
    }

    # Normal paragraph text
    Flush-List
    $processed = Fmt $line
    if ($paraBuf.Length -gt 0) { $paraBuf.Append(" $processed") | Out-Null }
    else                        { $paraBuf.Append($processed)   | Out-Null }
}

# Final flush
Flush-Para; Flush-List; Flush-Table
$parts.Add('</body></html>')

# Write HTML
$htmlText = $parts -join "`n"
$enc = New-Object System.Text.UTF8Encoding($false)   # UTF-8 без BOM
[System.IO.File]::WriteAllText($htmlPath, $htmlText, $enc)
Write-Host "HTML created: $htmlPath"

# ── Word: open HTML → save DOCX ──────────────────────────────────────────────
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Open($htmlPath, $false, $false, $false)

    # Page margins
    $doc.PageSetup.TopMargin    = $word.CentimetersToPoints(2)
    $doc.PageSetup.BottomMargin = $word.CentimetersToPoints(2)
    $doc.PageSetup.LeftMargin   = $word.CentimetersToPoints(2.5)
    $doc.PageSetup.RightMargin  = $word.CentimetersToPoints(2)

    # Save as DOCX (wdFormatDocumentDefault = 16)
    $fmt = [ref]16
    $path = [ref]$docPath
    $doc.SaveAs($path, $fmt)
    $doc.Close($false)
    Write-Host "DOCX saved: $docPath"
} catch {
    Write-Host "Error: $_"
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}

# Cleanup temp HTML
Remove-Item $htmlPath -Force -ErrorAction SilentlyContinue
Write-Host "Done."
