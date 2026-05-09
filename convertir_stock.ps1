# convertir_stock.ps1 — Genera stock.js desde Reporte Stock.csv
# Uso: powershell -ExecutionPolicy Bypass -File convertir_stock.ps1
# Ejecutar cada vez que se actualice Reporte Stock.csv

$csv = "C:\Users\usuario\Documents\Tableros\Reporte Stock.csv"
$out = "C:\Users\usuario\Documents\Tableros\stock.js"

if (-not (Test-Path $csv)) {
    Write-Host "ERROR: No se encontro el archivo: $csv" -ForegroundColor Red
    Read-Host "Presiona Enter para salir"
    exit 1
}

$lines = [System.IO.File]::ReadAllLines($csv, [System.Text.Encoding]::UTF8)

if ($lines.Length -lt 2) {
    Write-Host "ERROR: El archivo CSV esta vacio o no tiene datos." -ForegroundColor Red
    Read-Host "Presiona Enter para salir"
    exit 1
}

# Detectar separador
$sep = if ($lines[0].Contains(";")) { ";" } else { "," }

# Leer encabezados y detectar columnas
$headers = $lines[0] -split [regex]::Escape($sep) | ForEach-Object { $_.Trim().ToLower() -replace '"','' }

function Find-Col([string[]]$candidates) {
    foreach ($cand in $candidates) {
        for ($i = 0; $i -lt $headers.Count; $i++) {
            if ($headers[$i] -like "*$cand*") { return $i }
        }
    }
    return -1
}

$colProv      = Find-Col @("nombre","proveedor","marca")
$colArt       = Find-Col @("descripci","articulo","artículo","producto")
$colColorCod  = Find-Col @("color")          # primera coincidencia: "color"
$colTalle     = Find-Col @("talle","tamaño")
$colStock     = Find-Col @("cantidad","stock","disponible","existencia")

# "color descripción" suele estar antes que "color" en orden, buscar la segunda columna que contiene "color"
$colColorDesc = -1
for ($i = 0; $i -lt $headers.Count; $i++) {
    if ($headers[$i] -like "*color*" -and $i -ne $colColorCod) {
        $colColorDesc = $i
        break
    }
}

Write-Host "Columnas detectadas:" -ForegroundColor Cyan
Write-Host "  Proveedor        : $(if ($colProv     -ge 0) { $headers[$colProv]     + ' (col ' + $colProv     + ')' } else { 'NO ENCONTRADA' })"
Write-Host "  Articulo         : $(if ($colArt      -ge 0) { $headers[$colArt]      + ' (col ' + $colArt      + ')' } else { 'NO ENCONTRADA' })"
Write-Host "  Color codigo     : $(if ($colColorCod -ge 0) { $headers[$colColorCod] + ' (col ' + $colColorCod + ')' } else { 'NO ENCONTRADA' })"
Write-Host "  Color descripcion: $(if ($colColorDesc -ge 0) { $headers[$colColorDesc] + ' (col ' + $colColorDesc + ')' } else { 'NO ENCONTRADA' })"
Write-Host "  Talle            : $(if ($colTalle    -ge 0) { $headers[$colTalle]    + ' (col ' + $colTalle    + ')' } else { 'NO ENCONTRADA' })"
Write-Host "  Stock/Cantidad   : $(if ($colStock    -ge 0) { $headers[$colStock]    + ' (col ' + $colStock    + ')' } else { 'NO ENCONTRADA' })"

if ($colStock -lt 0) {
    Write-Host ""
    Write-Host "ERROR: No se encontro columna de stock/cantidad." -ForegroundColor Red
    Read-Host "Presiona Enter para salir"
    exit 1
}

$esc = [scriptblock]{ param($s) $s.Trim().Replace('\','\\').Replace('"','\"') }

$records = New-Object System.Collections.ArrayList
$errores = 0

for ($i = 1; $i -lt $lines.Length; $i++) {
    $line = $lines[$i].Trim() -replace '"',''
    if (-not $line) { continue }
    $c = $line -split [regex]::Escape($sep)

    $prov      = if ($colProv     -ge 0 -and $colProv     -lt $c.Length) { (& $esc $c[$colProv])     } else { "" }
    $art       = if ($colArt      -ge 0 -and $colArt      -lt $c.Length) { (& $esc $c[$colArt])      } else { "" }
    $colorCod  = if ($colColorCod -ge 0 -and $colColorCod -lt $c.Length) { (& $esc $c[$colColorCod]) } else { "" }
    $colorDesc = if ($colColorDesc -ge 0 -and $colColorDesc -lt $c.Length) { (& $esc $c[$colColorDesc]) } else { "" }
    $talle     = if ($colTalle    -ge 0 -and $colTalle    -lt $c.Length) { (& $esc $c[$colTalle])    } else { "" }
    $stkRaw    = if ($colStock    -ge 0 -and $colStock    -lt $c.Length) { $c[$colStock].Trim() } else { "0" }

    $stkVal = 0.0
    $parsed = [double]::TryParse(
        $stkRaw.Replace(",", "."),
        [System.Globalization.NumberStyles]::Any,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$stkVal
    )
    if (-not $parsed) { $errores++; continue }

    $r = '{"proveedor":"' + $prov + '","articulo":"' + $art + '","color":"' + $colorCod + '","color descripcion":"' + $colorDesc + '","talle":"' + $talle + '","stock":' + $stkVal.ToString([System.Globalization.CultureInfo]::InvariantCulture) + '}'
    [void]$records.Add($r)
}

$date    = Get-Date -Format "yyyy-MM-dd HH:mm"
$json    = "[" + ($records -join ",") + "]"
$content = "// stock.js - generado el $date desde Reporte Stock.csv`nwindow.STOCK_DATA=$json;"

[System.IO.File]::WriteAllText($out, $content, [System.Text.Encoding]::UTF8)

Write-Host ""
Write-Host "Listo: $out" -ForegroundColor Green
Write-Host "  Registros procesados : $($records.Count)"
if ($errores -gt 0) { Write-Host "  Filas con errores     : $errores" -ForegroundColor Yellow }
Write-Host ""
Write-Host "Recarga index.html en el navegador para ver el stock actualizado."
Read-Host "Presiona Enter para salir"
