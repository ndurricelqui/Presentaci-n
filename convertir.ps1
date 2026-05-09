# ============================================================
# convertir.ps1 — Genera datos.js para el Dashboard de Ventas
# Uso: powershell -ExecutionPolicy Bypass -File convertir.ps1
# Reemplazar Reporte Ventas.csv y volver a ejecutar para actualizar
# ============================================================

$csv = "C:\Users\usuario\Documents\Tableros\Reporte Ventas.csv"
$out = "C:\Users\usuario\Documents\Tableros\datos.js"

if (-not (Test-Path $csv)) {
    Write-Host "ERROR: No se encontro el archivo: $csv" -ForegroundColor Red
    exit 1
}

$lines = [System.IO.File]::ReadAllLines($csv, [System.Text.Encoding]::UTF8)
$records = New-Object System.Collections.ArrayList
$errores = 0

for ($i = 1; $i -lt $lines.Length; $i++) {
    $line = $lines[$i].Trim()
    if (-not $line) { continue }
    $c = $line -split ";"
    if ($c.Length -lt 12) { $errores++; continue }

    $qty = 0.0; $amt = 0.0
    [double]::TryParse($c[10].Trim().Replace(",","."), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$qty) | Out-Null
    [double]::TryParse($c[11].Trim().Replace(",","."), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$amt) | Out-Null

    $esc = [scriptblock]{ param($s) $s.Trim().Replace('\','\\').Replace('"','\"') }

    $r = '{"co":"'+(& $esc $c[0])+'","or":"'+(& $esc $c[1])+'","fe":"'+(& $esc $c[2])+'","ve":"'+(& $esc $c[3])+'","li":"'+(& $esc $c[4])+'","pr":"'+(& $esc $c[5])+'","ar":"'+(& $esc $c[6])+'","cd":"'+(& $esc $c[7])+'","cl":"'+(& $esc $c[8])+'","ta":"'+(& $esc $c[9])+'","ca":'+$qty.ToString([System.Globalization.CultureInfo]::InvariantCulture)+',"mo":'+$amt.ToString([System.Globalization.CultureInfo]::InvariantCulture)+'}'
    [void]$records.Add($r)
}

$json = "[" + ($records -join ",") + "]"
[System.IO.File]::WriteAllText($out, "window.VENTAS_DATA=$json;", [System.Text.Encoding]::UTF8)

Write-Host "Listo: $out" -ForegroundColor Green
Write-Host "  Registros procesados : $($records.Count)"
if ($errores -gt 0) { Write-Host "  Filas con errores     : $errores" -ForegroundColor Yellow }
Write-Host ""
Write-Host "Abri index.html en el navegador para ver el dashboard."
