# Excel Horizontal Capture v17
# Japanese UI / No Python required
# Windows PowerShell 5.1 + Desktop Microsoft Excel
# MIT License

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$ErrorActionPreference = "Stop"

Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class NativeWin {
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();

    [DllImport("user32.dll")]
    public static extern uint GetDpiForWindow(IntPtr hwnd);

    [DllImport("dwmapi.dll")]
    public static extern int DwmGetWindowAttribute(
        IntPtr hwnd,
        int dwAttribute,
        out RECT pvAttribute,
        int cbAttribute
    );
}
"@

try { [void][NativeWin]::SetProcessDPIAware() } catch {}

# Excel constants
$XL_MAXIMIZED = -4137

function Get-ExcelApp {
    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    } catch {
        throw "実行中の Excel を検出できませんでした。先に Excel ファイルを開いてください。"
    }
}

function Get-ExcelContext {
    $excel = Get-ExcelApp
    $wb = $excel.ActiveWorkbook
    $sheet = $excel.ActiveSheet
    $win = $excel.ActiveWindow

    if ($null -eq $wb -or $null -eq $sheet -or $null -eq $win) {
        throw "Excel にアクティブなブック、シート、またはウィンドウがありません。"
    }

    return [PSCustomObject]@{
        Excel = $excel
        Workbook = $wb
        Sheet = $sheet
        Window = $win
        WorkbookName = [string]$wb.Name
        SheetName = [string]$sheet.Name
    }
}

function Get-SelectionInfo {
    $ctx = Get-ExcelContext
    $sel = $ctx.Excel.Selection

    if ($null -eq $sel) {
        throw "Excel の現在の選択範囲を取得できませんでした。"
    }

    try {
        $addr = [string]$sel.Address($false, $false)
        $firstRow = [int]$sel.Row
        $firstCol = [int]$sel.Column
        $lastRow = $firstRow + [int]$sel.Rows.Count - 1
        $lastCol = $firstCol + [int]$sel.Columns.Count - 1
    } catch {
        throw "現在の選択対象は通常のセル範囲ではありません。マウスで連続したセル範囲を選択してください。"
    }

    return [PSCustomObject]@{
        Context = $ctx
        Address = $addr
        FirstRow = $firstRow
        FirstCol = $firstCol
        LastRow = $lastRow
        LastCol = $lastCol
    }
}

function Capture-ScreenRect([int]$x, [int]$y, [int]$w, [int]$h) {
    if ($w -le 0 -or $h -le 0) {
        throw "キャプチャ範囲が無効です：X=$x Y=$y W=$w H=$h"
    }

    $bmp = [System.Drawing.Bitmap]::new($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)

    try {
        $size = [System.Drawing.Size]::new($w, $h)
        $g.CopyFromScreen($x, $y, 0, 0, $size)
    } finally {
        $g.Dispose()
    }

    return $bmp
}

function Get-WindowVisualRect($ctx) {
    $hwnd = [IntPtr]::new([int64]$ctx.Excel.Hwnd)
    $rect = New-Object NativeWin+RECT

    try {
        $cb = [Runtime.InteropServices.Marshal]::SizeOf([type][NativeWin+RECT])
        $hr = [NativeWin]::DwmGetWindowAttribute($hwnd, 9, [ref]$rect, $cb)
        if ($hr -eq 0) {
            return $rect
        }
    } catch {}

    if (-not [NativeWin]::GetWindowRect($hwnd, [ref]$rect)) {
        throw "Excel メインウィンドウの位置を取得できませんでした。"
    }

    return $rect
}

function Get-SafeVisibleBottom($ctx) {
    $rect = Get-WindowVisualRect $ctx
    $hwnd = [IntPtr]::new([int64]$ctx.Excel.Hwnd)
    $screen = [System.Windows.Forms.Screen]::FromHandle($hwnd)
    $work = $screen.WorkingArea

    # Clamp to the Windows working area so the taskbar is never captured.
    return [Math]::Min([int]$rect.Bottom, [int]$work.Bottom)
}


function Get-WindowDpi($ctx) {
    try {
        $hwnd = [IntPtr]::new([int64]$ctx.Excel.Hwnd)
        $dpi = [int][NativeWin]::GetDpiForWindow($hwnd)
        if ($dpi -gt 0) { return $dpi }
    } catch {}
    return 96
}

function Get-ScaledPixels($ctx, [int]$basePixels) {
    $dpi = Get-WindowDpi $ctx
    return [int][Math]::Round($basePixels * $dpi / 96.0)
}

function Capture-TopFull($ctx, [int]$gridTopY, $originalTitleStrip) {
    $rect = Get-WindowVisualRect $ctx
    $h = $gridTopY - $rect.Top
    if ($h -lt (Get-ScaledPixels $ctx 70)) {
        $h = Get-ScaledPixels $ctx 70
    }

    $w = $rect.Right - $rect.Left
    $bmp = Capture-ScreenRect $rect.Left $rect.Top $w $h

    # Temporary workbook windows may display :2 in the title.
    Overlay-OriginalTitleStrip $bmp $originalTitleStrip 0
    return $bmp
}

function Capture-BottomFull($ctx) {
    $rect = Get-WindowVisualRect $ctx
    $safeBottom = Get-SafeVisibleBottom $ctx

    # Enough height for sheet tabs + horizontal scrollbar + status bar.
    # DPI-aware so 125% / 150% display scaling does not clip the tab row.
    $h = Get-ScaledPixels $ctx 86

    if ($h -gt ($safeBottom - $rect.Top)) {
        $h = [Math]::Max(1, $safeBottom - $rect.Top)
    }

    $top = $safeBottom - $h
    $w = $rect.Right - $rect.Left

    return Capture-ScreenRect $rect.Left $top $w $h
}

function Extend-BitmapRightEdge(
    [System.Drawing.Bitmap]$source,
    [int]$targetWidth
) {
    if ($source.Width -ge $targetWidth) {
        $copy = [System.Drawing.Bitmap]::new($source.Width, $source.Height)
        $g0 = [System.Drawing.Graphics]::FromImage($copy)
        try {
            $g0.DrawImageUnscaled($source, 0, 0)
        } finally {
            $g0.Dispose()
        }
        return $copy
    }

    $target = [System.Drawing.Bitmap]::new($targetWidth, $source.Height)
    $g = [System.Drawing.Graphics]::FromImage($target)

    try {
        $g.DrawImageUnscaled($source, 0, 0)

        # Extend only the final 1-pixel column of the original image.
        # This fills the otherwise-empty right side with the same Excel UI
        # background bands without repeating text/buttons.
        $srcRect = [System.Drawing.Rectangle]::new($source.Width - 1, 0, 1, $source.Height)
        $dstRect = [System.Drawing.Rectangle]::new(
            $source.Width,
            0,
            $targetWidth - $source.Width,
            $source.Height
        )
        $g.DrawImage($source, $dstRect, $srcRect, [System.Drawing.GraphicsUnit]::Pixel)
    } finally {
        $g.Dispose()
    }

    return $target
}

function Build-HeaderOnceFinal(
    [System.Drawing.Bitmap]$topBmp,
    [System.Drawing.Bitmap]$gridBmp,
    [System.Drawing.Bitmap]$bottomBmp,
    [bool]$includeBottom,
    [string]$savePath
) {
    $finalW = [Math]::Max($topBmp.Width, $gridBmp.Width)

    if ($includeBottom -and $bottomBmp -ne $null) {
        $finalW = [Math]::Max($finalW, $bottomBmp.Width)
    }

    $topWide = $null
    $bottomWide = $null
    $final = $null
    $g = $null

    try {
        $topWide = Extend-BitmapRightEdge $topBmp $finalW

        if ($includeBottom -and $bottomBmp -ne $null) {
            $bottomWide = Extend-BitmapRightEdge $bottomBmp $finalW
        }

        $finalH = $topWide.Height + $gridBmp.Height
        if ($bottomWide -ne $null) {
            $finalH += $bottomWide.Height
        }

        if ($finalW -gt 60000 -or $finalH -gt 30000) {
            throw "最終画像のサイズが大きすぎます：${finalW}x${finalH}px。"
        }

        $final = [System.Drawing.Bitmap]::new($finalW, $finalH)
        $g = [System.Drawing.Graphics]::FromImage($final)
        $g.Clear([System.Drawing.Color]::White)

        $y = 0
        $g.DrawImageUnscaled($topWide, 0, $y)
        $y += $topWide.Height

        $g.DrawImageUnscaled($gridBmp, 0, $y)
        $y += $gridBmp.Height

        if ($bottomWide -ne $null) {
            $g.DrawImageUnscaled($bottomWide, 0, $y)
        }

        $final.Save($savePath, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
        if ($g -ne $null) { $g.Dispose() }
        if ($final -ne $null) { $final.Dispose() }
        if ($topWide -ne $null) { $topWide.Dispose() }
        if ($bottomWide -ne $null) { $bottomWide.Dispose() }
    }
}

function Activate-ExcelWindow($ctx, [bool]$forceMaximize) {
    try { $ctx.Window.Activate() | Out-Null } catch {}
    try { $ctx.Sheet.Activate() | Out-Null } catch {}

    try {
        if ($forceMaximize) {
            $ctx.Window.WindowState = $XL_MAXIMIZED
        }
    } catch {}

    try {
        $hwnd = [IntPtr]::new([int64]$ctx.Excel.Hwnd)
        if ($forceMaximize) {
            [void][NativeWin]::ShowWindow($hwnd, 3) # SW_MAXIMIZE
        } else {
            [void][NativeWin]::ShowWindow($hwnd, 5) # SW_SHOW
        }
        [void][NativeWin]::SetForegroundWindow($hwnd)
    } catch {}
}

function Capture-OriginalTitleStrip($ctx) {
    $rect = Get-WindowVisualRect $ctx
    $w = $rect.Right - $rect.Left
    $h = [Math]::Min(58, [Math]::Max(44, [int](($rect.Bottom - $rect.Top) * 0.055)))
    return Capture-ScreenRect $rect.Left $rect.Top $w $h
}

function Overlay-OriginalTitleStrip(
    [System.Drawing.Bitmap]$target,
    [System.Drawing.Bitmap]$originalTitle,
    [int]$sourceX
) {
    if ($null -eq $target -or $null -eq $originalTitle) { return }

    $g = [System.Drawing.Graphics]::FromImage($target)
    try {
        $h = [Math]::Min($target.Height, $originalTitle.Height)
        $available = $originalTitle.Width - $sourceX
        if ($available -le 0) { return }

        $w = [Math]::Min($target.Width, $available)
        $src = [System.Drawing.Rectangle]::new($sourceX, 0, $w, $h)
        $dst = [System.Drawing.Rectangle]::new(0, 0, $w, $h)
        $g.DrawImage($originalTitle, $dst, $src, [System.Drawing.GraphicsUnit]::Pixel)
    } finally {
        $g.Dispose()
    }
}

function New-CaptureWindow($originalCtx) {
    $tempWin = $null

    try {
        $tempWin = $originalCtx.Workbook.NewWindow()
        $tempWin.Activate() | Out-Null

        # The capture window is always maximized for stable coordinates.
        try { $tempWin.WindowState = $XL_MAXIMIZED } catch {}
        try { $tempWin.Zoom = $originalCtx.Window.Zoom } catch {}
        try { $tempWin.DisplayGridlines = $originalCtx.Window.DisplayGridlines } catch {}
        try { $tempWin.DisplayHeadings = $true } catch {}
        try { $tempWin.DisplayWorkbookTabs = $true } catch {}
        try { $tempWin.DisplayHorizontalScrollBar = $true } catch {}
        try { $tempWin.DisplayVerticalScrollBar = $true } catch {}

        # Only the temporary window is modified.
        try { $tempWin.FreezePanes = $false } catch {}
        try { $tempWin.SplitColumn = 0 } catch {}
        try { $tempWin.SplitRow = 0 } catch {}
        try { $tempWin.SplitVertical = 0 } catch {}
        try { $tempWin.SplitHorizontal = 0 } catch {}

        try { $originalCtx.Sheet.Activate() | Out-Null } catch {}

        return [PSCustomObject]@{
            Excel = $originalCtx.Excel
            Workbook = $originalCtx.Workbook
            Sheet = $originalCtx.Sheet
            Window = $tempWin
            WorkbookName = $originalCtx.WorkbookName
            SheetName = $originalCtx.SheetName
        }
    } catch {
        if ($tempWin -ne $null) {
            try { $tempWin.Close() } catch {}
        }
        throw
    }
}

function Close-CaptureWindow($captureCtx, $originalCtx, [int]$originalWindowState, [bool]$wasSaved) {
    if ($captureCtx -ne $null) {
        try { $captureCtx.Window.Close() } catch {}
    }

    if ($originalCtx -ne $null) {
        try { $originalCtx.Window.Activate() | Out-Null } catch {}
        try { $originalCtx.Window.WindowState = $originalWindowState } catch {}

        # Creating/closing a view window can sometimes toggle the Saved flag.
        # If the workbook was clean before capture, restore only that clean flag.
        # If it already had unsaved edits, leave it untouched.
        if ($wasSaved) {
            try {
                if (-not [bool]$originalCtx.Workbook.Saved) {
                    $originalCtx.Workbook.Saved = $true
                }
            } catch {}
        }
    }
}

function Move-Selection-OutOfCaptureArea($selectionInfo, $captureCtx) {
    $sheet = $captureCtx.Sheet
    $helperRow = $null
    $helperCol = $selectionInfo.FirstCol

    if ($selectionInfo.LastRow -lt 1047576) {
        $helperRow = [Math]::Min(1048576, $selectionInfo.LastRow + 1000)
    } elseif ($selectionInfo.FirstRow -gt 1000) {
        $helperRow = [Math]::Max(1, $selectionInfo.FirstRow - 1000)
    } elseif ($selectionInfo.LastCol -lt 15384) {
        $helperRow = $selectionInfo.FirstRow
        $helperCol = [Math]::Min(16384, $selectionInfo.LastCol + 1000)
    } else {
        $helperRow = $selectionInfo.FirstRow
        $helperCol = $selectionInfo.FirstCol
    }

    try {
        $sheet.Cells.Item($helperRow, $helperCol).Select() | Out-Null
    } catch {}
}

function Get-CellScreenRect($ctx, [int]$row1, [int]$col1, [int]$row2, [int]$col2) {
    $pane = $ctx.Window.Panes.Item(1)
    $tl = $ctx.Sheet.Cells.Item($row1, $col1)

    if ($col2 -lt 16384) {
        $rightCell = $ctx.Sheet.Cells.Item($row1, $col2 + 1)
        $x2 = [int]$pane.PointsToScreenPixelsX([double]$rightCell.Left)
    } else {
        $lastCell = $ctx.Sheet.Cells.Item($row1, $col2)
        $x2 = [int]$pane.PointsToScreenPixelsX([double]$lastCell.Left + [double]$lastCell.Width)
    }

    if ($row2 -lt 1048576) {
        $bottomCell = $ctx.Sheet.Cells.Item($row2 + 1, $col1)
        $y2 = [int]$pane.PointsToScreenPixelsY([double]$bottomCell.Top)
    } else {
        $lastCell = $ctx.Sheet.Cells.Item($row2, $col1)
        $y2 = [int]$pane.PointsToScreenPixelsY([double]$lastCell.Top + [double]$lastCell.Height)
    }

    $x1 = [int]$pane.PointsToScreenPixelsX([double]$tl.Left)
    $y1 = [int]$pane.PointsToScreenPixelsY([double]$tl.Top)

    return [PSCustomObject]@{
        X1 = $x1
        Y1 = $y1
        X2 = $x2
        Y2 = $y2
        Width = $x2 - $x1
        Height = $y2 - $y1
    }
}

function Get-ColumnHeaderHeightPixels($ctx) {
    try {
        $pane = $ctx.Window.Panes.Item(1)
        $p0 = [int]$pane.PointsToScreenPixelsY(0.0)
        $p1 = [int]$pane.PointsToScreenPixelsY([double]$ctx.Sheet.StandardHeight)
        $rowPx = [Math]::Abs($p1 - $p0)

        # Excel column heading is slightly taller than a standard worksheet row.
        $headerPx = [int][Math]::Round($rowPx * 1.20)
        if ($headerPx -lt 22) { $headerPx = 22 }
        if ($headerPx -gt 50) { $headerPx = 50 }
        return $headerPx
    } catch {
        return 30
    }
}

function Get-GridSegmentRect(
    $ctx,
    [int]$row1,
    [int]$col1,
    [int]$row2,
    [int]$col2,
    [bool]$includeRowHeadings,
    [bool]$includeColumnHeadings
) {
    $cell = Get-CellScreenRect $ctx $row1 $col1 $row2 $col2
    $winRect = Get-WindowVisualRect $ctx

    $x1 = $cell.X1
    if ($includeRowHeadings) {
        $x1 = $winRect.Left + 2
    }

    $y1 = $cell.Y1
    if ($includeColumnHeadings) {
        $y1 = $cell.Y1 - (Get-ColumnHeaderHeightPixels $ctx)
    }

    # Keep away from the vertical scroll bar on the far right.
    $x2 = [Math]::Min($cell.X2, $winRect.Right - 26)

    return [PSCustomObject]@{
        X1 = $x1
        Y1 = $y1
        X2 = $x2
        Y2 = $cell.Y2
        Width = $x2 - $x1
        Height = $cell.Y2 - $y1
    }
}

function Get-FullyVisibleLastColumn($ctx, [int]$currentCol, [int]$targetLastCol) {
    $vr = $ctx.Window.VisibleRange
    $visibleLast = [int]$vr.Column + [int]$vr.Columns.Count - 1
    $candidate = [Math]::Min($targetLastCol, $visibleLast)

    if ($candidate -gt $currentCol -and $candidate -eq $visibleLast) {
        $candidate--
    }

    if ($candidate -lt $currentCol) {
        $candidate = $currentCol
    }

    $main = Get-WindowVisualRect $ctx

    while ($candidate -gt $currentCol) {
        $r = Get-CellScreenRect $ctx 1 $currentCol 1 $candidate
        if ($r.X2 -le ($main.Right - 26)) { break }
        $candidate--
    }

    return $candidate
}

function Get-FullyVisibleLastRow(
    $ctx,
    [int]$currentRow,
    [int]$targetLastRow,
    [int]$checkCol,
    [int]$bodyBottomLimit
) {
    $vr = $ctx.Window.VisibleRange
    $visibleLast = [int]$vr.Row + [int]$vr.Rows.Count - 1
    $candidate = [Math]::Min($targetLastRow, $visibleLast)

    if ($candidate -gt $currentRow -and $candidate -eq $visibleLast) {
        $candidate--
    }

    if ($candidate -lt $currentRow) {
        $candidate = $currentRow
    }

    while ($candidate -gt $currentRow) {
        $r = Get-CellScreenRect $ctx $currentRow $checkCol $candidate $checkCol
        if ($r.Y2 -le $bodyBottomLimit) { break }
        $candidate--
    }

    return $candidate
}

function Get-VisibleLastColumnFor2D(
    $ctx,
    [int]$currentCol,
    [int]$targetLastCol,
    [int]$checkRow
) {
    $vr = $ctx.Window.VisibleRange
    $visibleLast = [int]$vr.Column + [int]$vr.Columns.Count - 1
    $candidate = [Math]::Min($targetLastCol, $visibleLast)

    if ($candidate -gt $currentCol -and $candidate -eq $visibleLast) {
        $candidate--
    }

    if ($candidate -lt $currentCol) {
        $candidate = $currentCol
    }

    $main = Get-WindowVisualRect $ctx

    while ($candidate -gt $currentCol) {
        $r = Get-CellScreenRect $ctx $checkRow $currentCol $checkRow $candidate
        if ($r.X2 -le ($main.Right - 26)) { break }
        $candidate--
    }

    return $candidate
}

function Build-2DHeaderOnceFinal(
    [System.Drawing.Bitmap]$topBmp,
    [System.Drawing.Bitmap]$gridBmp,
    [System.Drawing.Bitmap]$bottomBmp,
    [bool]$includeBottom,
    [string]$savePath
) {
    $finalW = [Math]::Max($topBmp.Width, $gridBmp.Width)

    if ($includeBottom -and $bottomBmp -ne $null) {
        $finalW = [Math]::Max($finalW, $bottomBmp.Width)
    }

    $topWide = $null
    $bottomWide = $null
    $final = $null
    $g = $null

    try {
        $topWide = Extend-BitmapRightEdge $topBmp $finalW

        if ($includeBottom -and $bottomBmp -ne $null) {
            $bottomWide = Extend-BitmapRightEdge $bottomBmp $finalW
        }

        $finalH = $topWide.Height + $gridBmp.Height
        if ($bottomWide -ne $null) {
            $finalH += $bottomWide.Height
        }

        # Safety limits retained.
        if ($finalW -gt 60000) {
            throw "最終画像の横幅が上限を超えています：${finalW}px（上限 60000px）。"
        }
        if ($finalH -gt 60000) {
            throw "最終画像の縦幅が上限を超えています：${finalH}px（上限 60000px）。"
        }

        # Additional practical memory guard: 1.2 billion pixels ~= 4.8GB raw RGBA.
        $pixels = [int64]$finalW * [int64]$finalH
        if ($pixels -gt 300000000) {
            throw "最終画像が大きすぎます：${finalW}x${finalH}px。メモリ保護のため処理を停止しました。"
        }

        $final = [System.Drawing.Bitmap]::new($finalW, $finalH)
        $g = [System.Drawing.Graphics]::FromImage($final)
        $g.Clear([System.Drawing.Color]::White)

        $y = 0
        $g.DrawImageUnscaled($topWide, 0, $y)
        $y += $topWide.Height

        $g.DrawImageUnscaled($gridBmp, 0, $y)
        $y += $gridBmp.Height

        if ($bottomWide -ne $null) {
            $g.DrawImageUnscaled($bottomWide, 0, $y)
        }

        $final.Save($savePath, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
        if ($g -ne $null) { $g.Dispose() }
        if ($final -ne $null) { $final.Dispose() }
        if ($topWide -ne $null) { $topWide.Dispose() }
        if ($bottomWide -ne $null) { $bottomWide.Dispose() }
    }
}



function Capture-TopSegment($ctx, $gridRect, $originalTitleStrip) {
    $rect = Get-WindowVisualRect $ctx
    $topH = $gridRect.Y1 - $rect.Top

    if ($topH -lt 70) { $topH = 70 }
    if ($topH -gt ($rect.Bottom - $rect.Top - 80)) {
        throw "Excel 上部領域の計算に失敗しました。"
    }

    $x = $gridRect.X1
    $w = $gridRect.Width
    $bmp = Capture-ScreenRect $x $rect.Top $w $topH

    $sourceX = $x - $rect.Left
    Overlay-OriginalTitleStrip $bmp $originalTitleStrip $sourceX
    return $bmp
}

function Capture-BottomSegment($ctx, $gridRect) {
    $rect = Get-WindowVisualRect $ctx
    $safeBottom = Get-SafeVisibleBottom $ctx

    $h = Get-ScaledPixels $ctx 86
    $top = $safeBottom - $h

    if ($top -lt $rect.Top) {
        throw "Excel 下部領域の計算に失敗しました。"
    }

    $x = $gridRect.X1
    $w = [Math]::Min($gridRect.Width, ($rect.Right - 26) - $x)

    return Capture-ScreenRect $x $top $w $h
}

function Join-BitmapsHorizontally($bitmaps) {
    if ($bitmaps.Count -eq 0) {
        throw "結合する画像がありません。"
    }

    $totalW = 0
    $maxH = 0

    foreach ($bmp in $bitmaps) {
        if (-not ($bmp -is [System.Drawing.Image])) {
            throw "内部エラー：結合対象に画像以外のオブジェクトが含まれています。"
        }
        $totalW += [int]$bmp.Width
        if ($bmp.Height -gt $maxH) { $maxH = [int]$bmp.Height }
    }

    if ($totalW -gt 60000) {
        throw "横長画像が広すぎます（約 $totalW px）。1回にキャプチャする列数を減らしてください。"
    }

    if ($maxH -gt 30000) {
        throw "縦方向の画像が高すぎます（約 $maxH px）。"
    }

    $joined = [System.Drawing.Bitmap]::new($totalW, $maxH)
    $g = [System.Drawing.Graphics]::FromImage($joined)

    try {
        $g.Clear([System.Drawing.Color]::White)
        $x = 0

        foreach ($bmp in $bitmaps) {
            $g.DrawImageUnscaled([System.Drawing.Image]$bmp, $x, 0)
            $x += [int]$bmp.Width
        }
    } finally {
        $g.Dispose()
    }

    return $joined
}

function Join-BitmapsVertically($bitmaps) {
    if ($bitmaps.Count -eq 0) {
        throw "結合する画像がありません。"
    }

    $maxW = 0
    $totalH = 0

    foreach ($bmp in $bitmaps) {
        if (-not ($bmp -is [System.Drawing.Image])) {
            throw "内部エラー：結合対象に画像以外のオブジェクトが含まれています。"
        }
        $totalH += [int]$bmp.Height
        if ($bmp.Width -gt $maxW) { $maxW = [int]$bmp.Width }
    }

    if ($totalW -gt 60000) {
        throw "画像の横幅が広すぎます（約 $maxW px）。"
    }

    if ($totalH -gt 60000) {
        throw "縦長画像が高すぎます（約 $totalH px）。1回にキャプチャする行数を減らしてください。"
    }

    if ($maxW -gt 30000) {
        throw "画像の横幅が広すぎます（約 $maxW px）。"
    }

    $joined = [System.Drawing.Bitmap]::new($maxW, $totalH)
    $g = [System.Drawing.Graphics]::FromImage($joined)

    try {
        $g.Clear([System.Drawing.Color]::White)
        $y = 0

        foreach ($bmp in $bitmaps) {
            $g.DrawImageUnscaled([System.Drawing.Image]$bmp, 0, $y)
            $y += [int]$bmp.Height
        }
    } finally {
        $g.Dispose()
    }

    return $joined
}

function Build-VerticalSegment(
    [System.Drawing.Bitmap]$topBmp,
    [System.Drawing.Bitmap]$gridBmp,
    [System.Drawing.Bitmap]$bottomBmp,
    [bool]$includeTop,
    [bool]$includeBottom
) {
    $w = $gridBmp.Width
    $h = $gridBmp.Height

    if ($includeTop -and $topBmp -ne $null) {
        $w = [Math]::Max($w, $topBmp.Width)
        $h += $topBmp.Height
    }

    if ($includeBottom -and $bottomBmp -ne $null) {
        $w = [Math]::Max($w, $bottomBmp.Width)
        $h += $bottomBmp.Height
    }

    $segment = [System.Drawing.Bitmap]::new($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($segment)

    try {
        $g.Clear([System.Drawing.Color]::White)
        $y = 0

        if ($includeTop -and $topBmp -ne $null) {
            $g.DrawImageUnscaled($topBmp, 0, $y)
            $y += $topBmp.Height
        }

        $g.DrawImageUnscaled($gridBmp, 0, $y)
        $y += $gridBmp.Height

        if ($includeBottom -and $bottomBmp -ne $null) {
            $g.DrawImageUnscaled($bottomBmp, 0, $y)
        }
    } finally {
        $g.Dispose()
    }

    return $segment
}

function Save-Image([System.Drawing.Bitmap]$bmp, [string]$savePath) {
    $bmp.Save($savePath, [System.Drawing.Imaging.ImageFormat]::Png)
}

# ---------------- GUI ----------------

$form = [System.Windows.Forms.Form]::new()
$form.Text = "Excel 連続キャプチャツール v20"
$form.Size = [System.Drawing.Size]::new(820, 805)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$title = [System.Windows.Forms.Label]::new()
$title.Text = "Excel 横長キャプチャ / Horizontal Capture"
$title.Font = [System.Drawing.Font]::new("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
$title.Location = [System.Drawing.Point]::new(22, 18)
$title.AutoSize = $true
$form.Controls.Add($title)

$desc = [System.Windows.Forms.Label]::new()
$desc.Text = "横方向・縦方向・2D に対応。選んだ方向で使わない設定は自動的にグレーアウトします。`r`n2D は【上部1回 + 2D表 + 下部1回】で固定します。"
$desc.Location = [System.Drawing.Point]::new(24, 58)
$desc.Size = [System.Drawing.Size]::new(760, 46)
$form.Controls.Add($desc)


$directionGroup = [System.Windows.Forms.GroupBox]::new()
$directionGroup.Text = "スクロール方向"
$directionGroup.Location = [System.Drawing.Point]::new(24, 110)
$directionGroup.Size = [System.Drawing.Size]::new(760, 94)
$form.Controls.Add($directionGroup)

$radioHorizontal = [System.Windows.Forms.RadioButton]::new()
$radioHorizontal.Text = "横方向（左右にスクロールして結合）"
$radioHorizontal.Location = [System.Drawing.Point]::new(18, 27)
$radioHorizontal.Size = [System.Drawing.Size]::new(320, 24)
$radioHorizontal.Checked = $true
$directionGroup.Controls.Add($radioHorizontal)

$radioVertical = [System.Windows.Forms.RadioButton]::new()
$radioVertical.Text = "縦方向（上下にスクロールして結合）"
$radioVertical.Location = [System.Drawing.Point]::new(370, 27)
$radioVertical.Size = [System.Drawing.Size]::new(320, 24)
$directionGroup.Controls.Add($radioVertical)

$radio2D = [System.Windows.Forms.RadioButton]::new()
$radio2D.Text = "横 + 縦（2D 連続キャプチャ）"
$radio2D.Location = [System.Drawing.Point]::new(18, 54)
$radio2D.Size = [System.Drawing.Size]::new(320, 24)
$directionGroup.Controls.Add($radio2D)

$modeGroup = [System.Windows.Forms.GroupBox]::new()
$modeGroup.Text = "キャプチャモード"
$modeGroup.Location = [System.Drawing.Point]::new(24, 214)
$modeGroup.Size = [System.Drawing.Size]::new(760, 96)
$form.Controls.Add($modeGroup)

$radioHeaderOnce = [System.Windows.Forms.RadioButton]::new()
$radioHeaderOnce.Text = "Excel 上部を先頭に1回だけ付ける（連続モード）"
$radioHeaderOnce.Location = [System.Drawing.Point]::new(18, 26)
$radioHeaderOnce.Size = [System.Drawing.Size]::new(700, 24)
$radioHeaderOnce.Checked = $true
$modeGroup.Controls.Add($radioHeaderOnce)

$radioWithTop = [System.Windows.Forms.RadioButton]::new()
$radioWithTop.Text = "各スクロール位置ごとに Excel 上部を付けて結合する（セグメントモード）"
$radioWithTop.Location = [System.Drawing.Point]::new(18, 57)
$radioWithTop.Size = [System.Drawing.Size]::new(720, 24)
$modeGroup.Controls.Add($radioWithTop)

$appearanceCheck = [System.Windows.Forms.CheckBox]::new()
$appearanceCheck.Text = "選択範囲の強調表示を消して、通常の白背景でキャプチャする（推奨）"
$appearanceCheck.Location = [System.Drawing.Point]::new(24, 324)
$appearanceCheck.Size = [System.Drawing.Size]::new(740, 28)
$appearanceCheck.Checked = $true
$form.Controls.Add($appearanceCheck)

$rowNumberCheck = [System.Windows.Forms.CheckBox]::new()
$rowNumberCheck.Text = "横方向の2枚目以降にも行番号（1, 2, 3...）を含める"
$rowNumberCheck.Location = [System.Drawing.Point]::new(24, 356)
$rowNumberCheck.Size = [System.Drawing.Size]::new(360, 28)
$rowNumberCheck.Checked = $false
$form.Controls.Add($rowNumberCheck)

$rowNumberHint = [System.Windows.Forms.Label]::new()
$rowNumberHint.Text = "※ 使用できない方向では自動的に無効になります"
$rowNumberHint.Location = [System.Drawing.Point]::new(390, 359)
$rowNumberHint.Size = [System.Drawing.Size]::new(360, 24)
$form.Controls.Add($rowNumberHint)

$bottomCheck = [System.Windows.Forms.CheckBox]::new()
$bottomCheck.Text = "Excel 下部のシートタブ / ステータスバーを含める"
$bottomCheck.Location = [System.Drawing.Point]::new(24, 388)
$bottomCheck.Size = [System.Drawing.Size]::new(480, 28)
$bottomCheck.Checked = $true
$form.Controls.Add($bottomCheck)

$readBtn = [System.Windows.Forms.Button]::new()
$readBtn.Text = "現在の Excel 選択範囲を取得 / 更新"
$readBtn.Location = [System.Drawing.Point]::new(24, 430)
$readBtn.Size = [System.Drawing.Size]::new(255, 40)
$form.Controls.Add($readBtn)

$exportBtn = [System.Windows.Forms.Button]::new()
$exportBtn.Text = "スクロールして PNG を作成"
$exportBtn.Location = [System.Drawing.Point]::new(292, 430)
$exportBtn.Size = [System.Drawing.Size]::new(250, 40)
$form.Controls.Add($exportBtn)

$openBtn = [System.Windows.Forms.Button]::new()
$openBtn.Text = "出力フォルダーを開く"
$openBtn.Location = [System.Drawing.Point]::new(555, 430)
$openBtn.Size = [System.Drawing.Size]::new(229, 40)
$form.Controls.Add($openBtn)

$info = [System.Windows.Forms.TextBox]::new()
$info.Location = [System.Drawing.Point]::new(24, 488)
$info.Size = [System.Drawing.Size]::new(760, 118)
$info.Multiline = $true
$info.ReadOnly = $true
$form.Controls.Add($info)

$status = [System.Windows.Forms.Label]::new()
$status.Text = "状態：先に【現在の Excel 選択範囲を取得 / 更新】を押してください。"
$status.Location = [System.Drawing.Point]::new(24, 624)
$status.Size = [System.Drawing.Size]::new(760, 64)
$form.Controls.Add($status)

$script:LastOutputDir = [Environment]::GetFolderPath("Desktop")
$script:Selection = $null


function Update-OptionAvailability {
    $isH = $radioHorizontal.Checked
    $isV = $radioVertical.Checked
    $is2DMode = $radio2D.Checked

    # 2D always uses the continuous/header-once composition.
    if ($is2DMode) {
        $radioHeaderOnce.Checked = $true
        $modeGroup.Enabled = $false
    } else {
        $modeGroup.Enabled = $true
    }

    # This option only has meaning in:
    # Horizontal + segment mode.
    $rowOptionEnabled = ($isH -and $radioWithTop.Checked -and (-not $is2DMode))
    $rowNumberCheck.Enabled = $rowOptionEnabled
    $rowNumberHint.Enabled = $rowOptionEnabled

    if ($is2DMode) {
        $rowNumberHint.Text = "※ 2D：行番号は各行バンドの左端に自動表示"
    } elseif ($isV) {
        $rowNumberHint.Text = "※ 縦方向：行番号は連続性のため自動表示"
    } elseif ($radioHeaderOnce.Checked) {
        $rowNumberHint.Text = "※ 横方向・連続モード：行番号は左端に1回だけ表示"
    } else {
        $rowNumberHint.Text = "※ 横方向・セグメントモードでのみ変更できます"
    }
}

$radioHorizontal.Add_CheckedChanged({
    if ($radioHorizontal.Checked) { Update-OptionAvailability }
})

$radioVertical.Add_CheckedChanged({
    if ($radioVertical.Checked) { Update-OptionAvailability }
})

$radio2D.Add_CheckedChanged({
    if ($radio2D.Checked) { Update-OptionAvailability }
    else { Update-OptionAvailability }
})

$radioHeaderOnce.Add_CheckedChanged({
    Update-OptionAvailability
})

$radioWithTop.Add_CheckedChanged({
    Update-OptionAvailability
})

Update-OptionAvailability

$readBtn.Add_Click({
    try {
        # Bring Excel to the foreground before reading the selection.
        $ctx = Get-ExcelContext
        Activate-ExcelWindow $ctx $false
        Start-Sleep -Milliseconds 180

        $s = Get-SelectionInfo
        $script:Selection = $s

        $freezeText = if ($s.Context.Window.FreezePanes) { "あり（元画面は変更しません）" } else { "なし" }
        $stateText = if ([int]$s.Context.Window.WindowState -eq $XL_MAXIMIZED) { "最大化" } else { "通常 / その他" }

        $info.Text = "ファイル: $($s.Context.WorkbookName)`r`nシート: $($s.Context.SheetName)`r`n範囲: $($s.Address)`r`nウィンドウ枠の固定: $freezeText`r`nExcel ウィンドウ状態: $stateText"
        $status.Text = "状態：選択範囲を保存しました。次回の撮影でもこの範囲をそのまま使用できます。"

        $form.Activate()
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "確認", "OK", "Warning") | Out-Null
    }
})

$exportBtn.Add_Click({
    $gridParts = [System.Collections.ArrayList]::new()
    $bottomParts = [System.Collections.ArrayList]::new()
    $segmentParts = [System.Collections.ArrayList]::new()
    $rowStrips = [System.Collections.ArrayList]::new()

    $originalCtx = $null
    $captureCtx = $null
    $selectionInfo = $script:Selection
    $originalTitleStrip = $null

    $topOnce = $null
    $bottomOnce = $null
    $joinedGrid = $null
    $joinedBottom = $null
    $joinedSegments = $null
    $joined2D = $null

    $originalWindowState = $null
    $wasSaved = $false

    try {
        if ($selectionInfo -eq $null) {
            throw "先に【現在の Excel 選択範囲を取得 / 更新】を押して、キャプチャ範囲を登録してください。"
        }

        $isHorizontal = $radioHorizontal.Checked
        $isVertical = $radioVertical.Checked
        $is2D = $radio2D.Checked
        $isHeaderOnce = $radioHeaderOnce.Checked

        $originalCtx = $selectionInfo.Context
        $originalWindowState = [int]$originalCtx.Window.WindowState
        $wasSaved = [bool]$originalCtx.Workbook.Saved

        try {
            $null = $originalCtx.Workbook.Name
            $null = $originalCtx.Window.Caption
        } catch {
            $script:Selection = $null
            throw "登録済みの Excel ウィンドウが見つかりません。もう一度【現在の Excel 選択範囲を取得 / 更新】を押してください。"
        }

        $form.Hide()

        Activate-ExcelWindow $originalCtx $false
        Start-Sleep -Milliseconds 250
        $originalTitleStrip = Capture-OriginalTitleStrip $originalCtx

        $captureCtx = New-CaptureWindow $originalCtx
        Activate-ExcelWindow $captureCtx $true
        Start-Sleep -Milliseconds 500

        if ($appearanceCheck.Checked) {
            Move-Selection-OutOfCaptureArea $selectionInfo $captureCtx
        }

        $captureCtx.Window.ScrollRow = $selectionInfo.FirstRow
        $captureCtx.Window.ScrollColumn = $selectionInfo.FirstCol
        Activate-ExcelWindow $captureCtx $true
        Start-Sleep -Milliseconds 400

        $safeBottom = Get-SafeVisibleBottom $captureCtx
        $bottomReserve = if ($bottomCheck.Checked) { Get-ScaledPixels $captureCtx 92 } else { Get-ScaledPixels $captureCtx 8 }

        if ($isHorizontal) {
            $fitRect = Get-CellScreenRect `
                $captureCtx `
                $selectionInfo.FirstRow `
                $selectionInfo.FirstCol `
                $selectionInfo.LastRow `
                $selectionInfo.FirstCol

            if ($fitRect.Y2 -gt ($safeBottom - $bottomReserve)) {
                throw "選択した行数が現在の Excel 表示領域の高さを超えています。Excel の表示倍率を下げるか、選択行数を減らして再実行してください。"
            }

            $currentCol = $selectionInfo.FirstCol
            $segment = 0

            while ($currentCol -le $selectionInfo.LastCol) {
                $segment++

                $captureCtx.Window.ScrollColumn = $currentCol
                Activate-ExcelWindow $captureCtx $true
                Start-Sleep -Milliseconds 320

                $endCol = Get-FullyVisibleLastColumn $captureCtx $currentCol $selectionInfo.LastCol
                if ($endCol -lt $currentCol) { $endCol = $currentCol }

                if ($segment -eq 1) {
                    $includeRowHeadings = $true
                } elseif ((-not $isHeaderOnce) -and $rowNumberCheck.Checked) {
                    $includeRowHeadings = $true
                } else {
                    $includeRowHeadings = $false
                }

                $gridRect = Get-GridSegmentRect `
                    $captureCtx `
                    $selectionInfo.FirstRow `
                    $currentCol `
                    $selectionInfo.LastRow `
                    $endCol `
                    $includeRowHeadings `
                    $true

                if ($gridRect.Width -le 2 -or $gridRect.Height -le 2) {
                    throw "第 $segment セグメントのセル領域を計算できませんでした。"
                }

                $gridBmp = $null
                $bottomBmp = $null
                $topBmp = $null
                $segmentBmp = $null

                try {
                    $gridBmp = Capture-ScreenRect $gridRect.X1 $gridRect.Y1 $gridRect.Width $gridRect.Height

                    if ($isHeaderOnce) {
                        if ($segment -eq 1) {
                            $topOnce = Capture-TopFull $captureCtx $gridRect.Y1 $originalTitleStrip
                        }

                        [void]$gridParts.Add($gridBmp)
                        $gridBmp = $null

                        if ($bottomCheck.Checked) {
                            $bottomBmp = Capture-BottomSegment $captureCtx $gridRect
                            [void]$bottomParts.Add($bottomBmp)
                            $bottomBmp = $null
                        }

                    } else {
                        $topBmp = Capture-TopSegment $captureCtx $gridRect $originalTitleStrip

                        if ($bottomCheck.Checked) {
                            $bottomBmp = Capture-BottomSegment $captureCtx $gridRect
                        }

                        $segmentBmp = Build-VerticalSegment $topBmp $gridBmp $bottomBmp $true $bottomCheck.Checked
                        [void]$segmentParts.Add($segmentBmp)
                        $segmentBmp = $null
                    }
                } finally {
                    if ($gridBmp -ne $null) { try { $gridBmp.Dispose() } catch {} }
                    if ($bottomBmp -ne $null) { try { $bottomBmp.Dispose() } catch {} }
                    if ($topBmp -ne $null) { try { $topBmp.Dispose() } catch {} }
                    if ($segmentBmp -ne $null) { try { $segmentBmp.Dispose() } catch {} }
                }

                if ($endCol -ge $selectionInfo.LastCol) { break }
                $nextCol = $endCol + 1
                if ($nextCol -le $currentCol) {
                    throw "横スクロールが進まなかったため、無限ループ防止のため停止しました。"
                }
                $currentCol = $nextCol
            }

            if ($isHeaderOnce) {
                $joinedGrid = Join-BitmapsHorizontally $gridParts
                if ($bottomCheck.Checked) {
                    $joinedBottom = Join-BitmapsHorizontally $bottomParts
                }
            } else {
                $joinedSegments = Join-BitmapsHorizontally $segmentParts
            }

        } elseif ($isVertical) {
            $mainRect = Get-WindowVisualRect $captureCtx
            $fitRect = Get-CellScreenRect `
                $captureCtx `
                $selectionInfo.FirstRow `
                $selectionInfo.FirstCol `
                $selectionInfo.FirstRow `
                $selectionInfo.LastCol

            if ($fitRect.X2 -gt ($mainRect.Right - 26)) {
                throw "選択した列数が現在の Excel 表示領域の幅を超えています。Excel を最大化するか、表示倍率を下げるか、選択列数を減らして再実行してください。"
            }

            $currentRow = $selectionInfo.FirstRow
            $segment = 0

            while ($currentRow -le $selectionInfo.LastRow) {
                $segment++

                $captureCtx.Window.ScrollRow = $currentRow
                Activate-ExcelWindow $captureCtx $true
                Start-Sleep -Milliseconds 320

                $endRow = Get-FullyVisibleLastRow `
                    $captureCtx `
                    $currentRow `
                    $selectionInfo.LastRow `
                    $selectionInfo.FirstCol `
                    ($safeBottom - $bottomReserve)

                if ($endRow -lt $currentRow) { $endRow = $currentRow }

                $includeRowHeadings = $true
                $includeColumnHeadings = if ($isHeaderOnce) { $segment -eq 1 } else { $true }

                $gridRect = Get-GridSegmentRect `
                    $captureCtx `
                    $currentRow `
                    $selectionInfo.FirstCol `
                    $endRow `
                    $selectionInfo.LastCol `
                    $includeRowHeadings `
                    $includeColumnHeadings

                if ($gridRect.Width -le 2 -or $gridRect.Height -le 2) {
                    throw "第 $segment セグメントのセル領域を計算できませんでした。"
                }

                $gridBmp = $null
                $bottomBmp = $null
                $topBmp = $null
                $segmentBmp = $null

                try {
                    $gridBmp = Capture-ScreenRect $gridRect.X1 $gridRect.Y1 $gridRect.Width $gridRect.Height

                    if ($isHeaderOnce) {
                        if ($segment -eq 1) {
                            $topOnce = Capture-TopFull $captureCtx $gridRect.Y1 $originalTitleStrip
                            if ($bottomCheck.Checked) {
                                $bottomOnce = Capture-BottomFull $captureCtx
                            }
                        }

                        [void]$gridParts.Add($gridBmp)
                        $gridBmp = $null

                    } else {
                        $topBmp = Capture-TopSegment $captureCtx $gridRect $originalTitleStrip

                        if ($bottomCheck.Checked) {
                            $bottomBmp = Capture-BottomSegment $captureCtx $gridRect
                        }

                        $segmentBmp = Build-VerticalSegment $topBmp $gridBmp $bottomBmp $true $bottomCheck.Checked
                        [void]$segmentParts.Add($segmentBmp)
                        $segmentBmp = $null
                    }
                } finally {
                    if ($gridBmp -ne $null) { try { $gridBmp.Dispose() } catch {} }
                    if ($bottomBmp -ne $null) { try { $bottomBmp.Dispose() } catch {} }
                    if ($topBmp -ne $null) { try { $topBmp.Dispose() } catch {} }
                    if ($segmentBmp -ne $null) { try { $segmentBmp.Dispose() } catch {} }
                }

                if ($endRow -ge $selectionInfo.LastRow) { break }
                $nextRow = $endRow + 1
                if ($nextRow -le $currentRow) {
                    throw "縦スクロールが進まなかったため、無限ループ防止のため停止しました。"
                }
                $currentRow = $nextRow
            }

            if ($isHeaderOnce) {
                $joinedGrid = Join-BitmapsVertically $gridParts
            } else {
                $joinedSegments = Join-BitmapsVertically $segmentParts
            }

        } else {
            # 2D mode: split selection into tiles, stitch each row horizontally,
            # then stitch the completed row-strips vertically.
            #
            # For consistency, 2D uses header-once composition regardless of the
            # segment mode setting. This avoids repeating Excel chrome inside the grid.
            $captureCtx.Window.ScrollRow = $selectionInfo.FirstRow
            $captureCtx.Window.ScrollColumn = $selectionInfo.FirstCol
            Activate-ExcelWindow $captureCtx $true
            Start-Sleep -Milliseconds 350

            $firstGridRect = Get-GridSegmentRect `
                $captureCtx `
                $selectionInfo.FirstRow `
                $selectionInfo.FirstCol `
                $selectionInfo.FirstRow `
                $selectionInfo.FirstCol `
                $true `
                $true

            $topOnce = Capture-TopFull $captureCtx $firstGridRect.Y1 $originalTitleStrip

            if ($bottomCheck.Checked) {
                $bottomOnce = Capture-BottomFull $captureCtx
            }

            $currentRow = $selectionInfo.FirstRow
            $tileCount = 0
            $rowBand = 0

            while ($currentRow -le $selectionInfo.LastRow) {
                $rowBand++

                $captureCtx.Window.ScrollRow = $currentRow
                $captureCtx.Window.ScrollColumn = $selectionInfo.FirstCol
                Activate-ExcelWindow $captureCtx $true
                Start-Sleep -Milliseconds 320

                $endRow = Get-FullyVisibleLastRow `
                    $captureCtx `
                    $currentRow `
                    $selectionInfo.LastRow `
                    $selectionInfo.FirstCol `
                    ($safeBottom - $bottomReserve)

                if ($endRow -lt $currentRow) { $endRow = $currentRow }

                $rowTiles = [System.Collections.ArrayList]::new()
                $currentCol = $selectionInfo.FirstCol
                $colBand = 0

                try {
                    while ($currentCol -le $selectionInfo.LastCol) {
                        $colBand++
                        $tileCount++

                        $captureCtx.Window.ScrollColumn = $currentCol
                        Activate-ExcelWindow $captureCtx $true
                        Start-Sleep -Milliseconds 280

                        $endCol = Get-VisibleLastColumnFor2D `
                            $captureCtx `
                            $currentCol `
                            $selectionInfo.LastCol `
                            $currentRow

                        if ($endCol -lt $currentCol) { $endCol = $currentCol }

                        # In 2D continuous composition:
                        # - row numbers only on the leftmost tile of each row band
                        # - column headings only on the top row band
                        $includeRowHeadings = ($colBand -eq 1)
                        $includeColumnHeadings = ($rowBand -eq 1)

                        $gridRect = Get-GridSegmentRect `
                            $captureCtx `
                            $currentRow `
                            $currentCol `
                            $endRow `
                            $endCol `
                            $includeRowHeadings `
                            $includeColumnHeadings

                        if ($gridRect.Width -le 2 -or $gridRect.Height -le 2) {
                            throw "2D タイル ${rowBand}-${colBand} のセル領域を計算できませんでした。"
                        }

                        $tileBmp = Capture-ScreenRect `
                            $gridRect.X1 `
                            $gridRect.Y1 `
                            $gridRect.Width `
                            $gridRect.Height

                        [void]$rowTiles.Add($tileBmp)

                        if ($endCol -ge $selectionInfo.LastCol) { break }
                        $nextCol = $endCol + 1
                        if ($nextCol -le $currentCol) {
                            throw "2D 横スクロールが進まなかったため停止しました。"
                        }
                        $currentCol = $nextCol
                    }

                    $rowStrip = Join-BitmapsHorizontally $rowTiles
                    [void]$rowStrips.Add($rowStrip)

                } finally {
                    foreach ($bmp in $rowTiles) {
                        try { $bmp.Dispose() } catch {}
                    }
                }

                if ($endRow -ge $selectionInfo.LastRow) { break }

                $nextRow = $endRow + 1
                if ($nextRow -le $currentRow) {
                    throw "2D 縦スクロールが進まなかったため停止しました。"
                }
                $currentRow = $nextRow
            }

            $joined2D = Join-BitmapsVertically $rowStrips

            # Check final estimated dimensions before save.
            $estW = [Math]::Max($topOnce.Width, $joined2D.Width)
            $estH = $topOnce.Height + $joined2D.Height
            if ($bottomCheck.Checked -and $bottomOnce -ne $null) {
                $estH += $bottomOnce.Height
            }

            if ($estW -gt 60000) {
                throw "2D 最終画像の横幅が上限を超えています：${estW}px（上限 60000px）。"
            }
            if ($estH -gt 60000) {
                throw "2D 最終画像の縦幅が上限を超えています：${estH}px（上限 60000px）。"
            }

            $pixels = [int64]$estW * [int64]$estH
            if ($pixels -gt 300000000) {
                throw "2D 最終画像が大きすぎます：${estW}x${estH}px。メモリ保護のため処理を停止しました。"
            }

            $segment = $tileCount
        }

        Close-CaptureWindow $captureCtx $originalCtx $originalWindowState $wasSaved
        $captureCtx = $null

        $form.Show()
        $form.Activate()
        $status.Text = "状態：$segment セグメント / タイルを取得しました。保存先を選択してください。"
        $form.Refresh()

        $dlg = [System.Windows.Forms.SaveFileDialog]::new()
        $safeWb = [IO.Path]::GetFileNameWithoutExtension($originalCtx.WorkbookName)
        $dlg.Filter = "PNG image (*.png)|*.png"

        if ($isHorizontal) {
            if ($isHeaderOnce) {
                $dlg.FileName = "${safeWb}_$($originalCtx.SheetName)_horizontal_header_once.png"
            } else {
                $dlg.FileName = "${safeWb}_$($originalCtx.SheetName)_horizontal_segment.png"
            }
        } elseif ($isVertical) {
            if ($isHeaderOnce) {
                $dlg.FileName = "${safeWb}_$($originalCtx.SheetName)_vertical_header_once.png"
            } else {
                $dlg.FileName = "${safeWb}_$($originalCtx.SheetName)_vertical_segment.png"
            }
        } else {
            $dlg.FileName = "${safeWb}_$($originalCtx.SheetName)_2D.png"
        }

        $dlg.InitialDirectory = $script:LastOutputDir

        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            if ($is2D) {
                Build-2DHeaderOnceFinal `
                    $topOnce `
                    $joined2D `
                    $bottomOnce `
                    $bottomCheck.Checked `
                    $dlg.FileName

            } elseif ($isHeaderOnce) {
                if ($isHorizontal) {
                    Build-HeaderOnceFinal $topOnce $joinedGrid $joinedBottom $bottomCheck.Checked $dlg.FileName
                } else {
                    Build-HeaderOnceFinal $topOnce $joinedGrid $bottomOnce $bottomCheck.Checked $dlg.FileName
                }

            } else {
                Save-Image $joinedSegments $dlg.FileName
            }

            $script:LastOutputDir = Split-Path -Parent $dlg.FileName
            $status.Text = "状態：完了。範囲 $($selectionInfo.Address) は保持されています。続けてもう一度作成できます。"

            [System.Windows.Forms.MessageBox]::Show(
                "PNG を作成しました：`r`n$($dlg.FileName)",
                "完了",
                "OK",
                "Information"
            ) | Out-Null

        } else {
            $status.Text = "状態：保存をキャンセルしました。選択範囲は保持されています。"
        }

    } catch {
        try {
            if ($captureCtx -ne $null) {
                Close-CaptureWindow $captureCtx $originalCtx $originalWindowState $wasSaved
                $captureCtx = $null
            }
        } catch {}

        try {
            if (-not $form.Visible) {
                $form.Show()
                $form.Activate()
            }
        } catch {}

        $status.Text = "状態：失敗"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "エラー", "OK", "Error") | Out-Null

    } finally {
        foreach ($bmp in $gridParts) {
            try { $bmp.Dispose() } catch {}
        }
        foreach ($bmp in $bottomParts) {
            try { $bmp.Dispose() } catch {}
        }
        foreach ($bmp in $segmentParts) {
            try { $bmp.Dispose() } catch {}
        }
        foreach ($bmp in $rowStrips) {
            try { $bmp.Dispose() } catch {}
        }

        if ($originalTitleStrip -ne $null) { try { $originalTitleStrip.Dispose() } catch {} }
        if ($topOnce -ne $null) { try { $topOnce.Dispose() } catch {} }
        if ($bottomOnce -ne $null) { try { $bottomOnce.Dispose() } catch {} }
        if ($joinedGrid -ne $null) { try { $joinedGrid.Dispose() } catch {} }
        if ($joinedBottom -ne $null) { try { $joinedBottom.Dispose() } catch {} }
        if ($joinedSegments -ne $null) { try { $joinedSegments.Dispose() } catch {} }
        if ($joined2D -ne $null) { try { $joined2D.Dispose() } catch {} }

        try {
            if (-not $form.Visible) {
                $form.Show()
                $form.Activate()
            }
        } catch {}
    }
})
$openBtn.Add_Click({
    if (Test-Path $script:LastOutputDir) {
        Start-Process explorer.exe $script:LastOutputDir
    }
})

[void]$form.ShowDialog()
