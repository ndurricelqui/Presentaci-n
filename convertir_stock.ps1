# convertir_stock.ps1
# Lee C:\Users\usuario\Documents\PowerBi Pampero\Reporte Stock.xlsx
# Genera stock.js en el mismo directorio que este script
# Uso: clic derecho sobre el archivo → "Ejecutar con PowerShell"
#      o desde terminal: .\convertir_stock.ps1

$ExcelPath   = "C:\Users\usuario\Documents\PowerBi Pampero\Reporte Stock.xlsx"
$OutputPath  = Join-Path $PSScriptRoot "stock.js"

# ── Verificar que existe el archivo ──────────────────────────────────────────
if (-not (Test-Path $ExcelPath)) {
    Write-Host "ERROR: No se encontro el archivo:" -ForegroundColor Red
    Write-Host "  $ExcelPath" -ForegroundColor Red
    Write-Host ""
    Write-Host "Verifica la ruta y volvé a ejecutar." -ForegroundColor Yellow
    Read-Host "Presioná Enter para salir"
    exit 1
}

Write-Host "Leyendo: $ExcelPath" -ForegroundColor Cyan

# ── Abrir Excel via COM ───────────────────────────────────────────────────────
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($ExcelPath)
    $ws = $wb.Sheets.Item(1)   # Primera hoja
} catch {
    Write-Host "ERROR al abrir Excel: $_" -ForegroundColor Red
    Read-Host "Presioná Enter para salir"
    exit 1
}

# ── Leer cabecera y detectar columnas ─────────────────────────────────────────
$lastCol  = $ws.UsedRange.Columns.Count
$lastRow  = $ws.UsedRange.Rows.Count
$startRow = $ws.UsedRange.Row
$startCol = $ws.UsedRange.Column

Write-Host "Filas: $lastRow  |  Columnas: $lastCol" -ForegroundColor Gray

# Leer fila de encabezados
$headers = @()
for ($c = $startCol; $c -le ($startCol + $lastCol - 1); $c++) {
    $val = $ws.Cells.Item($startRow, $c).Text
    $headers += $val.ToString().Trim().ToLower()
}

function Find-Col($candidates) {
    foreach ($cand in $candidates) {
        for ($i = 0; $i -lt $headers.Count; $i++) {
            if ($headers[$i] -like "*$cand*") { return $i + $startCol }
        }
    }
    return -1
}

$colProv  = Find-Col @("proveedor","marca","nombre prov")
$colArt   = Find-Col @("articulo","artículo","descripcion","descripción","producto")
$colColor = Find-Col @("color")
$colTalle = Find-Col @("talle","tamaño")
$colStk   = Find-Col @("stock","disponible","existencia","cantidad")

Write-Host ""
Write-Host "Columnas detectadas:" -ForegroundColor Cyan
Write-Host "  Proveedor : $(if ($colProv  -ge 0) { $headers[$colProv  - $startCol] + ' (col ' + $colProv  + ')' } else { 'NO ENCONTRADA' })"
Write-Host "  Articulo  : $(if ($colArt   -ge 0) { $headers[$colArt   - $startCol] + ' (col ' + $colArt   + ')' } else { 'NO ENCONTRADA' })"
Write-Host "  Color     : $(if ($colColor -ge 0) { $headers[$colColor - $startCol] + ' (col ' + $colColor + ')' } else { 'NO ENCONTRADA' })"
Write-Host "  Talle     : $(if ($colTalle -ge 0) { $headers[$colTalle - $startCol] + ' (col ' + $colTalle + ')' } else { 'NO ENCONTRADA' })"
Write-Host "  Stock     : $(if ($colStk   -ge 0) { $headers[$colStk   - $startCol] + ' (col ' + $colStk   + ')' } else { 'NO ENCONTRADA' })"

if ($colStk -lt 0) {
    Write-Host ""
    Write-Host "ERROR: No se encontro columna de stock (stock/cantidad/disponible/existencia)." -ForegroundColor Red
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Read-Host "Presioná Enter para salir"
    exit 1
}

# ── Leer filas de datos ────────────────────────────────────────────────────────
$records = @()
$dataStart = $startRow + 1
$errors = 0

for ($row = $dataStart; $row -le ($startRow + $lastRow - 1); $row++) {
    $prov  = if ($colProv  -ge 0) { $ws.Cells.Item($row, $colProv ).Text.Trim() } else { "" }
    $art   = if ($colArt   -ge 0) { $ws.Cells.Item($row, $colArt  ).Text.Trim() } else { "" }
    $color = if ($colColor -ge 0) { $ws.Cells.Item($row, $colColor).Text.Trim() } else { "" }
    $talle = if ($colTalle -ge 0) { $ws.Cells.Item($row, $colTalle).Text.Trim() } else { "" }
    $stkRaw = $ws.Cells.Item($row, $colStk).Text.Trim().Replace(",", ".")

    # Saltar filas vacías
    if ($prov -eq "" -and $art -eq "" -and $stkRaw -eq "") { continue }

    $stkVal = 0.0
    $parsed = [double]::TryParse($stkRaw, [System.Globalization.NumberStyles]::Any,
                                  [System.Globalization.CultureInfo]::InvariantCulture, [ref]$stkVal)
    if (-not $parsed) {
        $errors++
        continue
    }

    # Escapar comillas para JSON
    $provEsc  = $prov  -replace '"', '\"'
    $artEsc   = $art   -replace '"', '\"'
    $colorEsc = $color -replace '"', '\"'
    $talleEsc = $talle -replace '"', '\"'

    $records += "  {""proveedor"":""$provEsc"",""articulo"":""$artEsc"",""color"":""$colorEsc"",""talle"":""$talleEsc"",""stock"":$stkVal}"
}

# ── Cerrar Excel ──────────────────────────────────────────────────────────────
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

# ── Escribir stock.js ─────────────────────────────────────────────────────────
$date   = Get-Date -Format "yyyy-MM-dd HH:mm"
$body   = $records -join ",`n"
$content = @"
// stock.js — generado automaticamente el $date
// Fuente: $ExcelPath
// NO editar manualmente. Ejecutar convertir_stock.ps1 para regenerar.
window.STOCK_DATA = [
$body
];
"@

[System.IO.File]::WriteAllText($OutputPath, $content, [System.Text.Encoding]::UTF8)

Write-Host ""
Write-Host "Listo!" -ForegroundColor Green
Write-Host "  Registros procesados : $($records.Count)" -ForegroundColor Green
if ($errors -gt 0) {
    Write-Host "  Filas con error de parseo (ignoradas): $errors" -ForegroundColor Yellow
}
Write-Host "  Archivo generado     : $OutputPath" -ForegroundColor Green
Write-Host ""
Write-Host "Copiá stock.js junto a index.html y recargá el tablero." -ForegroundColor Cyan
Write-Host ""
Read-Host "Presioná Enter para salir"
