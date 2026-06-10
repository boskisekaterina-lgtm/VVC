Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourcePath = Join-Path $root 'source_vvs_suppliers.jpg'
$source = [System.Drawing.Image]::FromFile($sourcePath)

$a3Width = 4961
$a3Height = 3508
$margin = 118
$printWidth = $a3Width - 2 * $margin
$printHeight = $a3Height - 2 * $margin

function New-A3Canvas {
    param([string]$Label)

    $bmp = New-Object System.Drawing.Bitmap($a3Width, $a3Height, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
    $bmp.SetResolution(300, 300)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::White)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality

    $font = New-Object System.Drawing.Font('Arial', 14, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(40, 48, 60))
    $g.DrawString($Label, $font, $brush, $margin, 48)
    $font.Dispose()
    $brush.Dispose()

    return @{ Bitmap = $bmp; Graphics = $g }
}

function Save-Png {
    param(
        [System.Drawing.Bitmap]$Bitmap,
        [string]$Path
    )

    $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/png' }
    $Bitmap.Save($Path, $codec, $null)
}

try {
    $pages = @()

    $overview = New-A3Canvas 'Лист 1: общий вид схемы на A3'
    $overviewG = $overview.Graphics
    $scale = [Math]::Min($printWidth / $source.Width, ($printHeight - 110) / $source.Height)
    $w = [int]($source.Width * $scale)
    $h = [int]($source.Height * $scale)
    $x = [int](($a3Width - $w) / 2)
    $y = [int](($a3Height - $h) / 2 + 20)
    $overviewG.DrawImage($source, $x, $y, $w, $h)
    $overviewPath = Join-Path $root '01_A3_overview.png'
    Save-Png $overview.Bitmap $overviewPath
    $pages += [pscustomobject]@{ File = '01_A3_overview.png'; Title = 'Лист 1. Общий вид' }
    $overviewG.Dispose()
    $overview.Bitmap.Dispose()

    $cropW = 560
    $cropH = 430
    $xStarts = @(0, 470, 940, ($source.Width - $cropW))
    $yStarts = @(0, ($source.Height - $cropH))
    $pageNo = 2

    for ($row = 0; $row -lt $yStarts.Count; $row++) {
        for ($col = 0; $col -lt $xStarts.Count; $col++) {
            $sx = [int]$xStarts[$col]
            $sy = [int]$yStarts[$row]
            $sw = [Math]::Min($cropW, $source.Width - $sx)
            $sh = [Math]::Min($cropH, $source.Height - $sy)
            $label = "Лист $pageNo`: увеличенный фрагмент $($row + 1)-$($col + 1), A3, текст ориентировочно 14 pt"
            $canvas = New-A3Canvas $label
            $g = $canvas.Graphics

            $scale = [Math]::Min($printWidth / $sw, ($printHeight - 110) / $sh)
            $dw = [int]($sw * $scale)
            $dh = [int]($sh * $scale)
            $dx = [int](($a3Width - $dw) / 2)
            $dy = [int](($a3Height - $dh) / 2 + 45)

            $srcRect = New-Object System.Drawing.Rectangle($sx, $sy, $sw, $sh)
            $dstRect = New-Object System.Drawing.Rectangle($dx, $dy, $dw, $dh)
            $g.DrawImage($source, $dstRect, $srcRect, [System.Drawing.GraphicsUnit]::Pixel)

            $name = '{0:D2}_A3_fragment_{1}_{2}.png' -f $pageNo, ($row + 1), ($col + 1)
            $path = Join-Path $root $name
            Save-Png $canvas.Bitmap $path
            $pages += [pscustomobject]@{ File = $name; Title = "Лист $pageNo. Фрагмент $($row + 1)-$($col + 1)" }

            $g.Dispose()
            $canvas.Bitmap.Dispose()
            $pageNo++
        }
    }

    $htmlPath = Join-Path $root 'vvs_suppliers_A3_print_layout.html'
    $pageHtml = ($pages | ForEach-Object {
        @"
  <section class="sheet">
    <img src="$($_.File)" alt="$($_.Title)">
  </section>
"@
    }) -join "`r`n"

    $html = @"
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <title>ВВС Поставщики - A3 печатный макет</title>
  <style>
    @page { size: A3 landscape; margin: 0; }
    * { box-sizing: border-box; }
    body { margin: 0; background: #e9edf2; }
    .sheet {
      width: 420mm;
      height: 297mm;
      page-break-after: always;
      break-after: page;
      background: white;
      overflow: hidden;
    }
    .sheet img {
      display: block;
      width: 420mm;
      height: 297mm;
    }
    @media screen {
      body { padding: 16px; }
      .sheet {
        margin: 0 auto 16px;
        box-shadow: 0 8px 32px rgba(20, 30, 45, .18);
      }
    }
  </style>
</head>
<body>
$pageHtml
</body>
</html>
"@
    Set-Content -LiteralPath $htmlPath -Value $html -Encoding UTF8
}
finally {
    $source.Dispose()
}
