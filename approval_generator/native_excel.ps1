param(
    [Parameter(Mandatory=$true)][string]$InputPath,
    [Parameter(Mandatory=$true)][string]$OutputPath,
    [Parameter(Mandatory=$true)][string]$DataPath,
    [string]$PdfPath = ''
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$excel = $null
$book = $null
try {
    if (Test-Path -LiteralPath $OutputPath) { throw 'Output already exists.' }
    $updates = Get-Content -LiteralPath $DataPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $excel.EnableEvents = $false
    $book = $excel.Workbooks.Open($InputPath, 0, $true, 5, '', '', $true)
    foreach ($group in $updates.PSObject.Properties) {
        $sheet = $book.Worksheets.Item($group.Name)
        try {
            $used = $sheet.UsedRange
            $lastOriginalRow = $used.Row + $used.Rows.Count - 1
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($used)
            $lastWrittenRow = $lastOriginalRow
            foreach ($update in $group.Value.PSObject.Properties) {
                $cell = $sheet.Range($update.Name)
                try {
                    if ($cell.HasFormula) { throw ('Formula target: ' + $group.Name + '!' + $update.Name) }
                    if ($cell.MergeCells) {
                        $anchor = $cell.MergeArea.Cells.Item(1, 1)
                        try {
                            if ($anchor.Row -ne $cell.Row -or $anchor.Column -ne $cell.Column) { throw ('Merged-cell anchor required: ' + $update.Name) }
                        } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($anchor) }
                    }
                    if ($null -eq $update.Value -or $update.Value -eq '') {
                        $cell.ClearContents() | Out-Null
                    } elseif ($update.Value -is [string]) {
                        # 텍스트로 저장한 뒤 기존 표시 형식을 복원하여 수식 실행을 방지한다.
                        $format = $cell.NumberFormat
                        $cell.NumberFormat = '@'
                        $cell.Value2 = [string]$update.Value
                        $cell.NumberFormat = $format
                        if ($cell.Row -gt $lastOriginalRow) { $cell.WrapText = $true }
                    } else {
                        $cell.Value2 = [double]$update.Value
                    }
                    if ($cell.Row -gt $lastWrittenRow) { $lastWrittenRow = $cell.Row }
                } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
            }
            if ($lastWrittenRow -gt $lastOriginalRow) {
                $newRows = $sheet.Range(([string]($lastOriginalRow + 1) + ':' + [string]$lastWrittenRow))
                $newRows.EntireRow.AutoFit() | Out-Null
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($newRows)
            }
        } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) }
    }
    $excel.CalculateFull()
    $book.CheckCompatibility = $false
    $book.SaveAs($OutputPath, 56)
    if ($PdfPath) { $book.ExportAsFixedFormat(0, $PdfPath) }
    Write-Output 'OK'
} catch {
    Write-Output $_.Exception.Message
    exit 1
} finally {
    if ($book) {
        $book.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    }
    if ($excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
