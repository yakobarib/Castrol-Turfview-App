# actualizar_bd.ps1
# ==================
# Actualiza los datos embebidos en la app cuando cambias las bases de datos.
#
# USO:
#   1. Reemplaza los archivos Excel en "3-Bases de Datos\"
#   2. Haz doble clic en este archivo (o ejecútalo en PowerShell)
#   3. Sube los cambios a GitHub automáticamente

param()

Set-Location $PSScriptRoot
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-CellValue($cell, $strings) {
    if ($cell.t -eq 's') { return $strings[[int]$cell.v] }
    return $cell.v
}
function Safe-String($val) { if ($null -eq $val) { return "" }; return [string]$val }

function Read-XlsxData($path) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
    $ssEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/sharedStrings.xml" }
    $strings = @()
    if ($ssEntry) {
        $sr = New-Object System.IO.StreamReader($ssEntry.Open())
        $xml = [xml]$sr.ReadToEnd(); $sr.Close()
        $strings = $xml.sst.si | ForEach-Object {
            $t = $_.t; if ($t -is [System.Xml.XmlElement]) { $t.'#text' } else { [string]$t }
        }
    }
    $sheetEntry = $zip.Entries | Where-Object { $_.FullName -match "xl/worksheets/sheet1\." }
    $sr2 = New-Object System.IO.StreamReader($sheetEntry.Open())
    $xml2 = [xml]$sr2.ReadToEnd(); $sr2.Close()
    $zip.Dispose()
    return @{ Rows = $xml2.worksheet.sheetData.row; Strings = $strings }
}

Write-Host "Leyendo BD_SKU_Codes.xlsx..."
$skuData = Read-XlsxData ".\3-Bases de Datos\BD_SKU_Codes.xlsx"
$skuDict = @{}
foreach ($row in $skuData.Rows) {
    $vals = @($row.c | ForEach-Object { Get-CellValue $_ $skuData.Strings })
    $sku = Safe-String $vals[0]
    $desc = Safe-String $vals[1]
    if ($sku -and $sku.ToUpper() -ne "SKU" -and $sku.ToUpper() -ne "SKU CODE") {
        $skuDict[$sku] = $desc
    }
}
Write-Host "  -> $($skuDict.Count) SKUs"

Write-Host "Leyendo BD_Precios_Tarifa.xlsx..."
$tarifaData = Read-XlsxData ".\3-Bases de Datos\BD_Precios_Tarifa.xlsx"
$catalogList = New-Object System.Collections.Specialized.OrderedDictionary
$dataStart = 0
$rows = @($tarifaData.Rows)
for ($i = 0; $i -lt [Math]::Min($rows.Count, 5); $i++) {
    $vals = @($rows[$i].c | ForEach-Object { Get-CellValue $_ $tarifaData.Strings })
    $r0 = Safe-String $vals[0]; $r1 = Safe-String $vals[1]
    if ($r0.ToUpper().Contains("SKU") -or $r1.ToUpper().Contains("PRODUCTO")) { $dataStart = $i + 1; break }
}
for ($i = $dataStart; $i -lt $rows.Count; $i++) {
    $vals = @($rows[$i].c | ForEach-Object { Get-CellValue $_ $tarifaData.Strings })
    $sku   = (Safe-String $vals[0]).Trim()
    $desc  = (Safe-String $vals[1]).Trim()
    $Lval  = Safe-String $vals[2]; $coste = Safe-String $vals[3]; $venta = Safe-String $vals[4]
    if (-not $sku) { continue }
    $Lf = 0.0; $cf = 0.0; $vf = 0.0
    $ic = [System.Globalization.CultureInfo]::InvariantCulture
    $ns = [System.Globalization.NumberStyles]::Any
    if (-not [double]::TryParse($Lval,$ns,$ic,[ref]$Lf)) { continue }
    if (-not [double]::TryParse($coste,$ns,$ic,[ref]$cf)) { continue }
    if (-not [double]::TryParse($venta,$ns,$ic,[ref]$vf)) { continue }
    if (-not $desc -and $skuDict.ContainsKey($sku)) { $desc = $skuDict[$sku] }
    $dJson = ConvertTo-Json $desc
    $skuJson = ConvertTo-Json $sku
    $entry = "{`"d`":$dJson,`"L`":$($Lf.ToString('G',$ic)),`"p`":$([Math]::Round($vf,6).ToString('G',$ic)),`"v`":$([Math]::Round($cf,6).ToString('G',$ic))}"
    $catalogList[$skuJson] = $entry
}

$pairs = @($catalogList.Keys | ForEach-Object { "$_`:$($catalogList[$_])" })
$json = "{" + ($pairs -join ",") + "}"
$count = $catalogList.Count
Write-Host "  -> $count referencias en catálogo"

Write-Host "Inyectando datos en el HTML..."
$htmlPath = ".\7- APP TURFVIEW\turfview_castrol_app_v2_3.html"
$html = [System.IO.File]::ReadAllText($htmlPath, [System.Text.Encoding]::UTF8)

$nuevo = "// <<<BD_EMBED_START>>>`r`n  const BD_EMBED_CATALOG = $json;`r`n  const BD_EMBED_COUNT   = $count;`r`n  // <<<BD_EMBED_END>>>"

if ($html.Contains("// <<<BD_EMBED_START>>>")) {
    $patron = [regex]"// <<<BD_EMBED_START>>>[\s\S]*?// <<<BD_EMBED_END>>>"
    $html = $patron.Replace($html, $nuevo)
    Write-Host "  -> Datos actualizados en el HTML"
} else {
    $pos = $html.IndexOf("// ─── BD DRAG")
    $html = $html.Substring(0, $pos) + $nuevo + "`r`n`r`n" + $html.Substring($pos)
    Write-Host "  -> Datos insertados en el HTML"
}
[System.IO.File]::WriteAllText($htmlPath, $html, [System.Text.Encoding]::UTF8)

Write-Host "Subiendo a GitHub..."
git add "7- APP TURFVIEW\turfview_castrol_app_v2_3.html" "3-Bases de Datos\BD_SKU_Codes.xlsx" "3-Bases de Datos\BD_Precios_Tarifa.xlsx"
git commit -m "feat: actualizar BD embebida ($count referencias)"
git push
Write-Host ""
Write-Host "Listo. La app en GitHub Pages tiene los datos actualizados."
Read-Host "Pulsa Enter para cerrar"
