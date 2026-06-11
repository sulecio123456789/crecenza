# migrar.ps1 - Crecensa: Referencias/*.xlsx -> datos-migrados.json

$baseDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$lotesXls = Join-Path $baseDir "Referencias\LOTES.xlsx"
$recXls   = Join-Path $baseDir "Referencias\RECIBOS CRECENSA.xlsx"
$outFile  = Join-Path $baseDir "datos-migrados.json"

foreach ($f in @($lotesXls, $recXls)) {
    if (-not (Test-Path $f)) { Write-Host "ERROR: No se encontro $f"; exit 1 }
}

function Parse-Precio($str) {
    if (-not $str) { return 0 }
    $clean = ($str -replace 'Q',''-replace ',',''-replace '\s',''-replace '-','0').Trim()
    if ($clean -eq '' -or $clean -eq '0') { return 0 }
    try { return [double]$clean } catch { return 0 }
}

function Map-Estado($pago) {
    $e = (($pago -replace '\s','').ToUpper())
    if ($e -match 'ESCRITURADO|CANCELADO')  { return 'vendida'   }
    if ($e -match 'PROCESO|RESERVADO')      { return 'reservado' }
    return 'disponible'
}

function Norm($s) {
    if (-not $s) { return '' }
    return ($s.Trim().ToUpper() -replace '\s+', ' ')
}

function Fix-Date($d) {
    if (-not $d -or $d.Trim() -eq '') { return '' }
    try { return ([datetime]::Parse($d.Trim())).ToString('yyyy-MM-dd') } catch { return '' }
}

function Cell($sheet, $row, $col) { return $sheet.Cells($row, $col).Text }

$script:cliId = 1
$script:terId = 1
$script:recId = 1
$clientes     = [System.Collections.Generic.List[object]]::new()
$terrenos     = [System.Collections.Generic.List[object]]::new()
$recibos      = [System.Collections.Generic.List[object]]::new()
$cliMap       = @{}

$skip = @('','ANULADO','N/A','SIN NOMBRE','PENDIENTE NOMBRE','PLANTA DE TRATAMIENTO','RESERVADO PROYECTOS','PROPIETARIO')

function Get-CliId($nombre, $tel) {
    $n = Norm $nombre
    if ($skip -contains $n) { return $null }
    if (-not $cliMap.ContainsKey($n)) {
        $id = $script:cliId++
        $cliMap[$n] = $id
        $clientes.Add([ordered]@{ id=$id; nom=$n; tel=''; dpi=''; vend=''; est='activo'; not='' })
    }
    $id = $cliMap[$n]
    if ($tel -and $tel.Trim() -ne '' -and $tel.Trim() -ne 'N/A') {
        $c = $clientes | Where-Object { $_.id -eq $id } | Select-Object -First 1
        if ($c -and $c.tel -eq '') { $c.tel = $tel.Trim() }
    }
    return $id
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # ----------------------------------------------------------------
    # 1. LOTES (Referencias\LOTES.xlsx  Hoja2)
    #    Col: [3]Sector [4]Lote [5]Fecha [6]Nombre [7]Precio [8]Reserva
    #         [9]Atraso [10]Pendiente [11]Pago(estado) [12]Plazo
    #         [13]Vendedor [14]NoRec [15]Telefono [16]Comentarios [2]Flag
    # ----------------------------------------------------------------
    Write-Host "Leyendo LOTES.xlsx..."
    $wb  = $excel.Workbooks.Open($lotesXls)
    $sh  = $wb.Sheets(1)
    $tot = $sh.UsedRange.Rows.Count

    for ($r = 2; $r -le $tot; $r++) {
        $sec = (Cell $sh $r 3).Trim().ToUpper()
        $lot = (Cell $sh $r 4).Trim().ToUpper()
        if (-not $sec -or -not $lot) { continue }

        $nombre   = Cell $sh $r 6
        $telefono = Cell $sh $r 15
        $cId      = Get-CliId $nombre $telefono

        $flag  = (Cell $sh $r 2).Trim()
        $nota  = (Cell $sh $r 16).Trim()
        if ($flag -and $flag -ne '') { $nota = if ($nota) {"$flag - $nota"} else {$flag} }

        $terrenos.Add([ordered]@{
            id    = $script:terId++
            sec   = $sec
            lot   = $lot
            pre   = (Parse-Precio (Cell $sh $r 7))
            sal   = (Parse-Precio (Cell $sh $r 10))
            est   = (Map-Estado   (Cell $sh $r 11))
            cliId = $cId
            vend  = (Norm (Cell $sh $r 13))
            are   = ''
            proj  = 'llanos'
            not   = $nota
        })
    }
    $wb.Close($false)
    Write-Host "  Lotes: $($terrenos.Count)"

    # ----------------------------------------------------------------
    # 2. RECIBOS 1-1000
    #    Col: [1]Fecha [2]NoRec [3]Cliente [4]Valor [5]Boleta
    #         [6]Lotes [7]Comentarios [8]Vendedor [9]ReciboOfi [10]FechaRep
    # ----------------------------------------------------------------
    Write-Host "Leyendo RECIBOS CRECENSA.xlsx  (1-1000)..."
    $wb2 = $excel.Workbooks.Open($recXls)

    $s1  = $wb2.Sheets("1-1000")
    $tot1 = $s1.UsedRange.Rows.Count
    for ($r = 2; $r -le $tot1; $r++) {
        $nom = Norm (Cell $s1 $r 3)
        if (-not $nom -or $nom -eq 'ANULADO') { continue }
        $mon = Parse-Precio (Cell $s1 $r 4)
        if ($mon -eq 0) { continue }

        $cId  = Get-CliId $nom ''
        $num  = (Cell $s1 $r 2).Trim(); if ($num -imatch '^n/a$') { $num = '' }
        $bol  = (Cell $s1 $r 5).Trim()
        $forma = 'Efectivo'
        if     ($bol -match '^\d')       { $forma = 'Deposito'      }
        elseif ($bol -imatch 'cheque')   { $forma = 'Cheque'        }
        elseif ($bol -imatch 'transfer') { $forma = 'Transferencia' }

        $lts  = (Cell $s1 $r 6).Trim()
        $com  = (Cell $s1 $r 7).Trim()
        $nota = ($lts + $(if ($com) {" - $com"} else {''})).Trim(' -')

        $recibos.Add([ordered]@{
            id    = $script:recId++
            fec   = (Fix-Date (Cell $s1 $r 1))
            fec2  = (Fix-Date (Cell $s1 $r 10))
            num   = $num
            cliId = $cId
            terId = $null
            mon   = $mon
            bol   = $forma
            vend  = (Norm (Cell $s1 $r 8))
            not   = $nota
        })
    }
    Write-Host "  Recibos 1-1000: $($recibos.Count)"

    # ----------------------------------------------------------------
    # 3. RECIBOS 1001-2000  (columna Egreso en [5], todo corre +1)
    #    Col: [1]Fecha [2]NoRec [3]Cliente [4]ValorIngreso [5]Egreso
    #         [6]Boleta [7]Lotes [8]Comentarios [9]Vendedor [10]ReciboOfi [11]FechaRep
    # ----------------------------------------------------------------
    Write-Host "Leyendo RECIBOS CRECENSA.xlsx  (1001-2000)..."
    $s2   = $wb2.Sheets("1001-2000")
    $tot2 = $s2.UsedRange.Rows.Count
    $antes = $recibos.Count
    for ($r = 2; $r -le $tot2; $r++) {
        $nom = Norm (Cell $s2 $r 3)
        if (-not $nom -or $nom -eq 'ANULADO') { continue }
        $mon = Parse-Precio (Cell $s2 $r 4)
        if ($mon -eq 0) { continue }

        $cId  = Get-CliId $nom ''
        $num  = (Cell $s2 $r 2).Trim(); if ($num -imatch '^n/a$') { $num = '' }
        $bol  = (Cell $s2 $r 6).Trim()
        $forma = 'Efectivo'
        if     ($bol -match '^\d')       { $forma = 'Deposito'      }
        elseif ($bol -imatch 'cheque')   { $forma = 'Cheque'        }
        elseif ($bol -imatch 'transfer') { $forma = 'Transferencia' }

        $lts  = (Cell $s2 $r 7).Trim()
        $com  = (Cell $s2 $r 8).Trim()
        $nota = ($lts + $(if ($com) {" - $com"} else {''})).Trim(' -')

        $recibos.Add([ordered]@{
            id    = $script:recId++
            fec   = (Fix-Date (Cell $s2 $r 1))
            fec2  = (Fix-Date (Cell $s2 $r 11))
            num   = $num
            cliId = $cId
            terId = $null
            mon   = $mon
            bol   = $forma
            vend  = (Norm (Cell $s2 $r 9))
            not   = $nota
        })
    }
    Write-Host "  Recibos 1001-2000: $($recibos.Count - $antes)"
    $wb2.Close($false)

} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host ""
Write-Host "=== RESUMEN ==="
Write-Host "  Clientes : $($clientes.Count)"
Write-Host "  Lotes    : $($terrenos.Count)"
Write-Host "  Recibos  : $($recibos.Count)"
Write-Host "==============="

$resultado = [ordered]@{
    generado  = (Get-Date -Format 'yyyy-MM-dd HH:mm')
    clientes  = @($clientes)
    terrenos  = @($terrenos)
    recibos   = @($recibos)
}
$json = $resultado | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText($outFile, $json, [System.Text.Encoding]::UTF8)
Write-Host ""
Write-Host "OK  Archivo generado: $outFile"
