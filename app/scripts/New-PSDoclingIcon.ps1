#Requires -Version 5.1
<#
.SYNOPSIS
    Generates app/DoclingFrontend/psdocling.ico
.DESCRIPTION
    Draws the PSDocling mark (blue rounded square, white "PD" monogram) at
    16, 24, 32, 48 and 256 px and writes a multi-size .ico with PNG frames.
    Re-run only when the mark changes; the .ico is committed.
#>
param(
    [string]$OutputPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'DoclingFrontend\psdocling.ico')
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

function New-Frame([int]$Size) {
    $bmp = New-Object System.Drawing.Bitmap $Size, $Size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.TextRenderingHint = 'AntiAliasGridFit'
    $g.Clear([System.Drawing.Color]::Transparent)

    $r = [Math]::Max(3, [int]($Size * 0.22))
    $rect = New-Object System.Drawing.Rectangle 0, 0, ($Size - 1), ($Size - 1)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc($rect.X, $rect.Y, $r * 2, $r * 2, 180, 90)
    $path.AddArc($rect.Right - $r * 2, $rect.Y, $r * 2, $r * 2, 270, 90)
    $path.AddArc($rect.Right - $r * 2, $rect.Bottom - $r * 2, $r * 2, $r * 2, 0, 90)
    $path.AddArc($rect.X, $rect.Bottom - $r * 2, $r * 2, $r * 2, 90, 90)
    $path.CloseFigure()

    $top = [System.Drawing.Color]::FromArgb(255, 77, 182, 255)
    $bottom = [System.Drawing.Color]::FromArgb(255, 3, 105, 161)
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush $rect, $top, $bottom, 90
    $g.FillPath($brush, $path)

    $text = if ($Size -le 16) { 'D' } else { 'PD' }
    $fontSize = if ($Size -le 16) { $Size * 0.72 } else { $Size * 0.42 }
    $font = New-Object System.Drawing.Font 'Segoe UI', $fontSize, ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Pixel)
    $format = New-Object System.Drawing.StringFormat
    $format.Alignment = 'Center'
    $format.LineAlignment = 'Center'
    $textRect = New-Object System.Drawing.RectangleF 0, ($Size * 0.02), $Size, $Size
    $g.DrawString($text, $font, [System.Drawing.Brushes]::White, $textRect, $format)

    $g.Dispose(); $brush.Dispose(); $font.Dispose(); $path.Dispose()
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    return , $ms.ToArray()
}

$sizes = 16, 24, 32, 48, 256
$frames = foreach ($s in $sizes) { , (New-Frame $s) }

$out = New-Object System.IO.MemoryStream
$w = New-Object System.IO.BinaryWriter $out
$w.Write([UInt16]0); $w.Write([UInt16]1); $w.Write([UInt16]$sizes.Count)
$offset = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) {
    $dim = if ($sizes[$i] -ge 256) { 0 } else { $sizes[$i] }
    $w.Write([byte]$dim); $w.Write([byte]$dim); $w.Write([byte]0); $w.Write([byte]0)
    $w.Write([UInt16]1); $w.Write([UInt16]32)
    $w.Write([UInt32]$frames[$i].Length); $w.Write([UInt32]$offset)
    $offset += $frames[$i].Length
}
foreach ($f in $frames) { $w.Write($f) }
$w.Flush()
[System.IO.File]::WriteAllBytes($OutputPath, $out.ToArray())
Write-Host "Wrote $OutputPath ($($out.Length) bytes, sizes $($sizes -join ', '))"
