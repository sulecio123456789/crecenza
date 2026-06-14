# migrar-supabase.ps1 v3 - CRECENZA: Excel -> Supabase (bigint IDs)
# Archivos: Referencias\LOTES CRECENSA ACTUALIZADO.xlsx + Referencias\RECIBOS CRECENSA.xlsx

$SB_URL = 'https://ppmqvelqaqqamuypbogu.supabase.co'
$SB_KEY = 'sb_publishable_zufUHTqPkGHQ2JP-DvzfpA_w7sli1lk'

$HDR_WRITE = @{
    'apikey'        = $SB_KEY
    'Authorization' = "Bearer $SB_KEY"
    'Content-Type'  = 'application/json'
    'Prefer'        = 'return=minimal'
}
$HDR_READ = @{
    'apikey'        = $SB_KEY
    'Authorization' = "Bearer $SB_KEY"
}

$baseDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$lotesXls = Join-Path $baseDir "Referencias\LOTES CRECENSA ACTUALIZADO.xlsx"
$recXls   = Join-Path $baseDir "Referencias\RECIBOS CRECENSA.xlsx"

foreach ($f in @($lotesXls, $recXls)) {
    if (-not (Test-Path $f)) { Write-Host "ERROR: No se encontro $f"; exit 1 }
}

# ---- Helpers ----

function Parse-Precio($str) {
    if (-not $str) { return 0 }
    $clean = ($str -replace 'Q','' -replace ',','' -replace '\s','').Trim() -replace '-','0'
    if ($clean -eq '' -or $clean -eq '0') { return 0 }
    try { return [double]$clean } catch { return 0 }
}

function Map-Estado($s) {
    $e = (($s -replace '\s','').ToUpper())
    if ($e -match 'ESCRITURADO|CANCELADO') { return 'vendida'   }
    if ($e -match 'PROCESO|RESERVADO')     { return 'reservado' }
    return 'disponible'
}

function Norm($s) {
    if (-not $s) { return '' }
    return ($s.Trim().ToUpper() -replace '\s+', ' ')
}

function Fix-Date($d) {
    if (-not $d -or $d.Trim() -eq '' -or $d.Trim() -eq 'N/A') { return $null }
    $s = $d.Trim()
    if ($s -match '^(\d{1,2})/(\d{1,2})/(\d{4})$') {
        try { return ([datetime]::new([int]$Matches[3], [int]$Matches[2], [int]$Matches[1])).ToString('yyyy-MM-dd') }
        catch { return $null }
    }
    if ($s -match '^\d{4}-\d{2}-\d{2}$') { return $s }
    try { return ([datetime]::Parse($s)).ToString('yyyy-MM-dd') } catch { return $null }
}

function Parse-DiaPago($str) {
    if (-not $str) { return 0 }
    $m = [regex]::Match($str, '\d+')
    if ($m.Success) { return [int]$m.Value } else { return 0 }
}

function Cell($sh, $r, $c) { return $sh.Cells($r, $c).Text }

function Build-Distrib($lotsStr, $monTotal) {
    if (-not $lotsStr -or $lotsStr.Trim() -eq '' -or $lotsStr.Trim() -eq 'N/A') { return 'null' }
    $matches2 = [regex]::Matches($lotsStr, '([A-Za-z]+)[/\s]*(\d+)')
    if ($matches2.Count -eq 0) { return 'null' }
    $terIds = [System.Collections.Generic.List[int]]::new()
    foreach ($m in $matches2) {
        $sec = $m.Groups[1].Value.ToUpper().Trim()
        $lot = $m.Groups[2].Value.Trim()
        $key = $sec + '/' + $lot
        if ($terLookup.ContainsKey($key)) { $terIds.Add($terLookup[$key]) }
    }
    if ($terIds.Count -eq 0) { return 'null' }
    $montoXlote = [Math]::Round($monTotal / $terIds.Count, 2)
    $items = $terIds | ForEach-Object { '{"terId":' + $_ + ',"monto":' + $montoXlote + '}' }
    return '[' + ($items -join ',') + ']'
}

function EscJson($s) {
    if ($null -eq $s) { return 'null' }
    $s2 = $s.ToString() -replace '\\','\\' -replace '"','\"' -replace "`r",' ' -replace "`n",' '
    return '"' + $s2 + '"'
}

function SB-ErrBody($ex) {
    try {
        $stream = $ex.Exception.Response.GetResponseStream()
        return (New-Object System.IO.StreamReader($stream)).ReadToEnd()
    } catch { return $ex.Exception.Message }
}

function SB-Delete($table) {
    $url = "$SB_URL/rest/v1/$table" + '?id=gte.1'
    try {
        Invoke-RestMethod -Uri $url -Method Delete -Headers $HDR_READ | Out-Null
        Write-Host ("  DELETE $table OK")
    } catch {
        Write-Host ("  DELETE $table (vacia o error): " + (SB-ErrBody $_))
    }
}

function SB-Insert($table, $jsonRows) {
    if ($jsonRows.Count -eq 0) { Write-Host ("  (0 en $table, skip)"); return }
    $batchSize = 100
    $total = $jsonRows.Count
    $sent  = 0
    for ($i = 0; $i -lt $total; $i += $batchSize) {
        $end   = [Math]::Min($i + $batchSize - 1, $total - 1)
        $batch = $jsonRows[$i..$end]
        $json  = '[' + ($batch -join ',') + ']'
        try {
            Invoke-RestMethod -Uri "$SB_URL/rest/v1/$table" -Method Post `
                -Headers $HDR_WRITE `
                -Body ([System.Text.Encoding]::UTF8.GetBytes($json)) | Out-Null
            $sent += $batch.Count
            Write-Host ("  " + $sent + "/" + $total + " " + $table)
        } catch {
            Write-Host ("  ERROR batch $i : " + (SB-ErrBody $_))
        }
    }
}

# ---- Datos ----

$script:cliId = 1
$script:terId = 1
$script:recId = 1

$clientes = [System.Collections.Generic.List[string]]::new()
$terrenos = [System.Collections.Generic.List[string]]::new()
$recibos  = [System.Collections.Generic.List[string]]::new()
$cliMap   = @{}
$terLookup = @{}  # "SEC/LOT" -> terId (para distribucion de recibos)

# Contadores para meta-stats
$meta = @{ h1_validos=0; h1_anulado=0; h1_monto0=0; h1_vacios=0
           h2_validos=0; h2_anulado=0; h2_monto0=0; h2_vacios=0 }

$skipNames = @('','ANULADO','N/A','SIN NOMBRE','PENDIENTE NOMBRE','PLANTA DE TRATAMIENTO',
               'RESERVADO PROYECTOS','PROPIETARIO','PENDIENTE DATOS')

function Get-CliId($nombre, $tel) {
    $n = Norm $nombre
    if (-not $n) { return 'null' }
    foreach ($sk in $skipNames) { if ($n -eq $sk) { return 'null' } }
    if ($n -match '^PENDIENTE|^RESERVADO|^SIN NOMBRE|^PROPIETARIO|^ANULADO') { return 'null' }

    if (-not $cliMap.ContainsKey($n)) {
        $id = $script:cliId++
        $cliMap[$n] = $id
        $telClean = ''
        if ($tel -and $tel.Trim() -ne '' -and $tel.Trim() -ne 'N/A') { $telClean = $tel.Trim() }
        $clientes.Add((
            '{"id":' + $id +
            ',"nom":'  + (EscJson $n) +
            ',"tel":'  + (EscJson $telClean) +
            ',"dpi":""' +
            ',"vend":""' +
            ',"est":"activo"' +
            ',"nota":""}'
        ))
    }
    return $cliMap[$n]
}

# ---- Excel ----

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # ============================================================
    # LOTES CRECENSA ACTUALIZADO.xlsx
    # ============================================================
    Write-Host "Leyendo LOTES CRECENSA ACTUALIZADO.xlsx..."
    $wb = $excel.Workbooks.Open($lotesXls)
    $sh = $wb.Sheets(1)
    $tot = $sh.UsedRange.Rows.Count

    for ($r = 2; $r -le $tot; $r++) {
        $sec = (Cell $sh $r 1).Trim().ToUpper()
        $lot = (Cell $sh $r 2).Trim().ToUpper()
        if (-not $sec -or -not $lot) { continue }

        $cId      = Get-CliId (Cell $sh $r 4) (Cell $sh $r 15)
        $estadoCol = (Cell $sh $r 10).Trim()
        $plStr     = (Cell $sh $r 11).Trim()
        $plazo     = 0
        if ($plStr -match '^(\d+)') { try { $plazo = [int]$Matches[1] } catch {} }
        $nota  = (Cell $sh $r 16).Trim()
        $err20 = (Cell $sh $r 20).Trim()
        if ($err20) { $nota = if ($nota) { "$err20 - $nota" } else { $err20 } }
        $fecha = Fix-Date (Cell $sh $r 3)
        $fechaJ = if ($fecha) { '"' + $fecha + '"' } else { 'null' }

        $terLookup[$sec + '/' + $lot] = $script:terId
        $terrenos.Add((
            '{"id":'           + $script:terId +
            ',"sec":'          + (EscJson $sec) +
            ',"lot":'          + (EscJson $lot) +
            ',"pre":'          + (Parse-Precio (Cell $sh $r 5)) +
            ',"sal":'          + (Parse-Precio (Cell $sh $r 7)) +
            ',"est":"'         + (Map-Estado $estadoCol) + '"' +
            ',"cli_id":'       + $cId +
            ',"vend":'         + (EscJson (Norm (Cell $sh $r 12))) +
            ',"are":""' +
            ',"proj":"nichos"' +
            ',"nota":'         + (EscJson $nota) +
            ',"fecha_contrato":' + $fechaJ +
            ',"reserva":'      + (Parse-Precio (Cell $sh $r 6)) +
            ',"plazo":'        + $plazo +
            ',"dia_pago":'     + (Parse-DiaPago (Cell $sh $r 8)) +
            ',"comision":'     + (EscJson (Cell $sh $r 13).Trim()) +
            ',"estado_escrit":' + (EscJson $estadoCol) +
            '}'
        ))
        $script:terId++
    }
    $wb.Close($false)
    Write-Host ("  Lotes: " + $terrenos.Count)

    # ============================================================
    # RECIBOS - Sheet 1-1000
    # ============================================================
    Write-Host "Leyendo RECIBOS CRECENSA.xlsx (1-1000)..."
    $wb2  = $excel.Workbooks.Open($recXls)
    $s1   = $wb2.Sheets("1-1000")
    $tot1 = $s1.UsedRange.Rows.Count

    for ($r = 2; $r -le $tot1; $r++) {
        $nom = Norm (Cell $s1 $r 3)
        if (-not $nom) { $meta.h1_vacios++; continue }
        if ($nom -eq 'ANULADO') { $meta.h1_anulado++; continue }
        $mon = Parse-Precio (Cell $s1 $r 4)
        if ($mon -eq 0) { $meta.h1_monto0++; continue }
        $cId = Get-CliId $nom ''
        if ($cId -eq 'null') { $meta.h1_vacios++; continue }
        $meta.h1_validos++

        $num = (Cell $s1 $r 2).Trim(); if ($num -imatch '^n/a$') { $num = '' }
        $bol = (Cell $s1 $r 5).Trim()
        $forma = 'Efectivo'
        if     ($bol -match '^\d')       { $forma = 'Deposito'      }
        elseif ($bol -imatch 'cheque')   { $forma = 'Cheque'        }
        elseif ($bol -imatch 'transfer') { $forma = 'Transferencia' }
        $lts  = (Cell $s1 $r 6).Trim()
        $com  = (Cell $s1 $r 7).Trim()
        $nota = ($lts + $(if ($com) { " - $com" } else { '' })).Trim(' -')
        $fec  = Fix-Date (Cell $s1 $r 1)
        $fec2 = Fix-Date (Cell $s1 $r 10)
        $fecJ  = if ($fec)  { '"' + $fec  + '"' } else { 'null' }
        $fec2J = if ($fec2) { '"' + $fec2 + '"' } else { 'null' }
        $distribJ = Build-Distrib $lts $mon

        $recibos.Add((
            '{"id":'       + $script:recId +
            ',"fec":'      + $fecJ +
            ',"fec2":'     + $fec2J +
            ',"num":'      + (EscJson $num) +
            ',"cli_id":'   + $cId +
            ',"ter_id":null' +
            ',"mon":'      + $mon +
            ',"bol":'      + (EscJson $forma) +
            ',"vend":'     + (EscJson (Norm (Cell $s1 $r 8))) +
            ',"nota":'     + (EscJson $nota) +
            ',"distrib":'  + $distribJ +
            '}'
        ))
        $script:recId++
    }
    Write-Host ("  Recibos 1-1000: " + $recibos.Count)

    # ============================================================
    # RECIBOS - Sheet 1001-2000
    # ============================================================
    Write-Host "Leyendo RECIBOS CRECENSA.xlsx (1001-2000)..."
    $s2    = $wb2.Sheets("1001-2000")
    $tot2  = $s2.UsedRange.Rows.Count
    $antes = $recibos.Count

    for ($r = 2; $r -le $tot2; $r++) {
        $nom = Norm (Cell $s2 $r 3)
        if (-not $nom) { $meta.h2_vacios++; continue }
        if ($nom -eq 'ANULADO') { $meta.h2_anulado++; continue }
        $mon = Parse-Precio (Cell $s2 $r 4)
        if ($mon -eq 0) { $meta.h2_monto0++; continue }
        $cId = Get-CliId $nom ''
        if ($cId -eq 'null') { $meta.h2_vacios++; continue }
        $meta.h2_validos++

        $num = (Cell $s2 $r 2).Trim(); if ($num -imatch '^n/a$') { $num = '' }
        $bol = (Cell $s2 $r 6).Trim()
        $forma = 'Efectivo'
        if     ($bol -match '^\d')       { $forma = 'Deposito'      }
        elseif ($bol -imatch 'cheque')   { $forma = 'Cheque'        }
        elseif ($bol -imatch 'transfer') { $forma = 'Transferencia' }
        $lts  = (Cell $s2 $r 7).Trim()
        $com  = (Cell $s2 $r 8).Trim()
        $nota = ($lts + $(if ($com) { " - $com" } else { '' })).Trim(' -')
        $fec  = Fix-Date (Cell $s2 $r 1)
        $fec2 = Fix-Date (Cell $s2 $r 11)
        $fecJ  = if ($fec)  { '"' + $fec  + '"' } else { 'null' }
        $fec2J = if ($fec2) { '"' + $fec2 + '"' } else { 'null' }
        $distribJ = Build-Distrib $lts $mon

        $recibos.Add((
            '{"id":'       + $script:recId +
            ',"fec":'      + $fecJ +
            ',"fec2":'     + $fec2J +
            ',"num":'      + (EscJson $num) +
            ',"cli_id":'   + $cId +
            ',"ter_id":null' +
            ',"mon":'      + $mon +
            ',"bol":'      + (EscJson $forma) +
            ',"vend":'     + (EscJson (Norm (Cell $s2 $r 9))) +
            ',"nota":'     + (EscJson $nota) +
            ',"distrib":'  + $distribJ +
            '}'
        ))
        $script:recId++
    }
    Write-Host ("  Recibos 1001-2000: " + ($recibos.Count - $antes))
    $wb2.Close($false)

} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host ""
Write-Host "=== RESUMEN ==="
Write-Host ("  Clientes : " + $clientes.Count)
Write-Host ("  Lotes    : " + $terrenos.Count)
Write-Host ("  Recibos  : " + $recibos.Count)
Write-Host "==============="
Write-Host ""
Write-Host "Limpiando tablas..."
SB-Delete 'recibos'
SB-Delete 'terrenos'
SB-Delete 'clientes'
Write-Host ""
Write-Host "Insertando clientes..."
SB-Insert 'clientes' $clientes
Write-Host "Insertando lotes..."
SB-Insert 'terrenos' $terrenos
Write-Host "Insertando recibos..."
SB-Insert 'recibos' $recibos

# Guardar meta-stats de migracion
Write-Host "Guardando estadisticas de migracion..."
$fecha = (Get-Date -Format 'yyyy-MM-dd')
$metaJson = '[' + (
    @(
        "{`"key`":`"mig_fecha`",`"val`":`"$fecha`"}",
        "{`"key`":`"mig_h1_validos`",`"val`":`"$($meta.h1_validos)`"}",
        "{`"key`":`"mig_h1_anulado`",`"val`":`"$($meta.h1_anulado)`"}",
        "{`"key`":`"mig_h1_monto0`",`"val`":`"$($meta.h1_monto0)`"}",
        "{`"key`":`"mig_h1_vacios`",`"val`":`"$($meta.h1_vacios)`"}",
        "{`"key`":`"mig_h2_validos`",`"val`":`"$($meta.h2_validos)`"}",
        "{`"key`":`"mig_h2_anulado`",`"val`":`"$($meta.h2_anulado)`"}",
        "{`"key`":`"mig_h2_monto0`",`"val`":`"$($meta.h2_monto0)`"}",
        "{`"key`":`"mig_h2_vacios`",`"val`":`"$($meta.h2_vacios)`"}"
    ) -join ','
) + ']'
$hdrUpsert = @{
    'apikey'        = $SB_KEY
    'Authorization' = "Bearer $SB_KEY"
    'Content-Type'  = 'application/json'
    'Prefer'        = 'resolution=merge-duplicates,return=minimal'
}
try {
    Invoke-RestMethod -Uri "$SB_URL/rest/v1/meta" -Method Post `
        -Headers $hdrUpsert -Body ([System.Text.Encoding]::UTF8.GetBytes($metaJson)) | Out-Null
    Write-Host "  Meta-stats guardados OK"
} catch {
    Write-Host ("  Meta-stats (tabla no existe aun, crear en Supabase SQL): " + (SB-ErrBody $_))
}
Write-Host ""
Write-Host "MIGRACION COMPLETADA. Recarga la app en el navegador."
