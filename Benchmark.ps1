<#
================================================================================
 BENCHMARK LUZ CNMC - VERSION POWERSHELL (sin Python)
 - 25 empresas: API CNMC (datos reales)
 - 3 empresas:  Web scraping (Holaluz, Podo, Factor Energia)
 - Excel con historico semanal + comparativa, guardado en Descargas
 Requisitos: PowerShell 5.1 + Microsoft Excel instalado. NO necesita Python.
================================================================================
#>

# ------------------------------------------------------------------ CONFIG -----
$CodigoPostal    = 28003
$Potencia        = 4
$ConsumoAnualE   = 210
$Tarifa          = 4        # peaje 2.0TD domestico (NO cambiar a 1: el API devuelve 0 ofertas)
$RevisionPrecios = 1        # 1 = precio fijo (CNMC). Cambiar a 2 si se desea el otro conjunto
$CarpetaSalida   = "$env:USERPROFILE\Downloads"
$NombreExcel     = "Analisis_Energia_CNMC.xlsx"

$BaseUrl = 'https://comparador.cnmc.gob.es/api/publico'
$UA      = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'

# 25 empresas del API: NOMBRE_LEGAL -> @{ brand=marca; search=cadena de oferta }
$TargetOffers = [ordered]@{
    'ENERGYA VM'         = @{ brand='Energya VM';      search='Formula Fija Unica' }
    'DOMESTICA GAS'      = @{ brand='Visalia';         search='Visalia Luz Fijo' }
    'NATURGY'            = @{ brand='Naturgy';         search='Por Uso Luz' }
    'REPSOL'             = @{ brand='Repsol';          search='Tarifa Ahorro Potencia' }
    'OCTOPUS'            = @{ brand='Octopus';         search='OCTOPUS RELAX' }
    'NIBA'               = @{ brand='Niba';            search='niba Zen' }
    'ENERGIA NUFRI'      = @{ brand='Energia Nufri';   search='CALMA' }
    'GAOLANIA'           = @{ brand='Gana Energia';    search='Tarifa 24 horas' }
    'CLEARVIEW'          = @{ brand='Clarity Energy';  search='CLARITY ENERGY' }
    'CIDE'               = @{ brand='CHC Energia';     search='Ilumina' }
    'ENERGYASSET'        = @{ brand='Energy Asset';    search='Tarifa Plana' }
    'CATGAS'             = @{ brand='Catgas';          search='2.0TDL' }
    'TELECOR'            = @{ brand='El Corte Ingles'; search='Despreocupate' }
    'ENDESA'             = @{ brand='Endesa';          search='Luz Fija 24h' }
    'FENIE ENERGIA'      = @{ brand='Fenie Energia';   search='Fijo Energetico' }
    'GESTERNOVA'         = @{ brand='Contigo Energia'; search='Tarifa Facil' }
    'DISA ENERGIA'       = @{ brand='Disa Energia';    search='ALISIOS' }
    'HIDROELECTRICA'     = @{ brand='HSC';             search='Eficiente' }
    'LUMISA'             = @{ brand='Lumisa Energia';  search='2.0TD' }
    'TOTALENERGIES'      = @{ brand='Total Energies';  search='TU AIRE' }
    'WEKIWI'             = @{ brand='Wekiki';          search='MariCalmen' }
    'PLENITUDE'          = @{ brand='Plenitude';       search='Tarifa Facil Plus' }
    'IMAGINA'            = @{ brand='Imagina';         search='PLAN BASE' }
    'IBERDROLA'          = @{ brand='Iberdrola';       search='Plan Online' }
    'NEXUS'              = @{ brand='Nexus';           search='Luz Eficiente' }
    # Estas 3 se intentan primero en CNMC (como benchmark2); si no estan, caen a web scraping
    'HOLALUZ'            = @{ brand='Holaluz';         search='Tarifa Clasica' }
    'GEO ALTERNATIVA'    = @{ brand='Podo';            search='Luz Precio Unico 24h' }
    'FACTORENERGIA'      = @{ brand='Factor Energia';  search='Tarifa Unica' }
}

# ----------------------------------------------------- TLS + certificados ------
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
if (-not ([System.Management.Automation.PSTypeName]'CnmcTrustAll').Type) {
    Add-Type @"
using System.Net; using System.Security.Cryptography.X509Certificates;
public class CnmcTrustAll : ICertificatePolicy {
  public bool CheckValidationResult(ServicePoint s, X509Certificate c, WebRequest r, int p) { return true; }
}
"@
}
[System.Net.ServicePointManager]::CertificatePolicy = New-Object CnmcTrustAll

# --------------------------------------------------------------- HELPERS -------

function Get-FullParams {
    param([int]$IdOferta = 0)
    $fin    = (Get-Date -Day 1).AddDays(-1)        # ultimo dia del mes anterior
    $inicio = Get-Date -Date $fin -Day 1           # primer dia del mes anterior
    $tsIni  = [DateTimeOffset]::new([DateTime]::SpecifyKind($inicio,'Utc')).ToUnixTimeMilliseconds()
    $tsFin  = [DateTimeOffset]::new([DateTime]::SpecifyKind($fin,'Utc')).ToUnixTimeMilliseconds()

    [ordered]@{
        tipoSuministro='E'; codigoPostal=$CodigoPostal; potencia=$Potencia
        potenciaPrimeraFranja=$Potencia; potenciaSegundaFranja=$Potencia; potenciaTerceraFranja=$Potencia
        potenciaCuartaFranja=$Potencia; potenciaQuintaFranja=$Potencia; potenciaSextaFranja=$Potencia
        consumoAnualE=$ConsumoAnualE; consumoAnualEOrig=2600; consumoPrimeraFranja=61; consumoSegundaFranja=51
        consumoTerceraFranja=98; consumoCuartaFranja=0; consumoQuintaFranja=0; consumoSextaFranja=0
        consumoAnualEQr=0; consumoPrimeraFranjaQr=0; consumoSegundaFranjaQr=0; consumoTerceraFranjaQr=0
        consumoCuartaFranjaQr=0; consumoQuintaFranjaQr=0; consumoSextaFranjaQr=0; consumoAnualEPQr=0
        consumoPrimeraFranjaPQr=0; consumoSegundaFranjaPQr=0; consumoTerceraFranjaPQr=0; consumoCuartaFranjaPQr=0
        consumoQuintaFranjaPQr=0; consumoSextaFranjaPQr=0; tarifa=$Tarifa; consumoAnualG=491; consumoAnualGOrig=6000
        serviciosAdicionales=2; permanencia=2; idOferta=$IdOferta; vivienda='true'; factura='true'
        energiaAutoconsumo=0; idAuditoriaQR=0; potenciaAutoconsumo=3.5; revisionPrecios=$RevisionPrecios; importe=0
        dateInicio=$tsIni; dateFin=$tsFin; tc=0; bs=0; impSA=0; impOtros=0; exc=0; reg=0; mecanismoAjuste=0
        importeMecanismoAjustePunta=0; importeMecanismoAjusteLlano=0; importeMecanismoAjusteValle=0
        precioConsumoMecanismoAjustePunta=0; precioConsumoMecanismoAjusteLlano=0; precioConsumoMecanismoAjusteValle=0
        precioConsumoMecanismoAjusteTotal=0; mecanismoAjusteIVA=0; impOtrosConIE=0; impOtrosSinIE=0
        pmaxP1=0; pmaxP2=0; fFact=$tsFin; dtoBS=0; finBS=0; ajuste=0; impPot=0; impEner=0; dto=0
        prP1=0; prP2=0; prE1=0; prE2=0; prE3=0; cfP1flex=0; cfP2flex=0; cambio=0; promo=0; verde=0
        rev=0; trampeo=0; perfilConsumo=13; cups='0000'; autoconsumo='false'
    }
}

function ConvertTo-QueryString {
    param([System.Collections.IDictionary]$Params)
    ($Params.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '&'
}

function Invoke-CnmcApi {
    param([string]$Url)
    $resp = Invoke-WebRequest -Uri $Url -Headers @{ 'User-Agent' = $UA } -TimeoutSec 60 -UseBasicParsing
    $json = [System.Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())  # arregla UTF-8
    $json | ConvertFrom-Json
}

function Normalize-Name {
    # Quita acentos y pasa a minusculas, para casar nombres entre semanas
    # (p.ej. "El Corte Ingles" == "El Corte Ingles" con tilde).
    param([string]$Text)
    if (-not $Text) { return '' }
    $d = ([string]$Text).Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $d.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    return $sb.ToString().Trim().ToLower()
}

function Get-Number {
    param([string]$Text)
    if (-not $Text) { return $null }
    try { return [double]::Parse(($Text -replace ',', '.'), [Globalization.CultureInfo]::InvariantCulture) }
    catch { return $null }
}

function Find-InRange {
    <# Devuelve el primer numero (segun patrones) que caiga en [min,max]. #>
    param([string]$Zona, [string[]]$Patrones, [double]$Min, [double]$Max)
    foreach ($p in $Patrones) {
        foreach ($m in [regex]::Matches($Zona, $p, 'IgnoreCase')) {
            $v = Get-Number $m.Groups[1].Value
            if ($v -ne $null -and $v -ge $Min -and $v -le $Max) { return $v }
        }
    }
    return $null
}

function Parse-Prices {
    <#
      Devuelve @{ energia=...; p1=...; p2=... } a partir del texto 'caracteristicas'.
      Estrategia por SECCIONES: el texto varia mucho por comercializadora y a veces la
      seccion de energia va antes que la de potencia (p.ej. Nexus). Por eso primero se
      separan la zona de POTENCIA y la de ENERGIA por sus encabezados y luego se busca
      dentro de cada una. '.' en los patrones sustituye a letras acentuadas (e/i/n).
    #>
    param([string]$Texto)
    $r = @{ energia = $null; p1 = $null; p2 = $null }
    if (-not $Texto) { return $r }
    $t = $Texto -replace [char]0x00A0, ' '   # nbsp -> espacio

    # --- Localizar encabezados de cada seccion ---
    # El lookahead exige un PRECIO (numero decimal) dentro de 32 caracteres tras el
    # encabezado, para ignorar frases introductorias tipo "dos precios de potencia y un
    # precio de energia" (cuyo primer decimal queda mas lejos).
    $reHdrPot = '(?:precios?\s+(?:del\s+)?(?:t.rmino\s+)?(?:de\s+)?potencia|t.rmino[s]?\s+(?:de\s+|de\s+la\s+)?potencia)(?=[\s\S]{0,32}[0-9]+[.,][0-9])'
    $reHdrEne = '(?:precios?\s+(?:del\s+)?(?:t.rmino\s+)?(?:de\s+)?energ.a|t.rmino[s]?\s+(?:de\s+|de\s+la\s+)?energ.a)(?=[\s\S]{0,32}[0-9]+[.,][0-9])'
    $mPot = [regex]::Match($t, $reHdrPot, 'IgnoreCase')
    $mEne = [regex]::Match($t, $reHdrEne, 'IgnoreCase')

    # --- Definir zonas (potencia / energia) segun el orden en que aparecen ---
    $zonaPot = $t; $zonaEne = $t
    if ($mPot.Success -and $mEne.Success) {
        if ($mPot.Index -lt $mEne.Index) {
            $zonaPot = $t.Substring($mPot.Index, $mEne.Index - $mPot.Index)
            $zonaEne = $t.Substring($mEne.Index)
        } else {
            $zonaEne = $t.Substring($mEne.Index, $mPot.Index - $mEne.Index)
            $zonaPot = $t.Substring($mPot.Index)
        }
    } elseif ($mPot.Success) {
        $zonaPot = $t.Substring($mPot.Index)
    } elseif ($mEne.Success) {
        $zonaEne = $t.Substring($mEne.Index)
    }

    # --- Energia (E/kWh) 0.05-0.30: preferir precio SIN descuento ---
    # '.' (punto) no matchea salto de linea, asi que '.*?' se limita a la misma linea.
    $r.energia = Find-InRange $zonaEne @(
        'sin\s+descuentos?.*?([0-9]+[.,][0-9]{3,6})',
        '([0-9]+[.,][0-9]{3,6})\s*[^0-9]{0,4}kWh',
        '([0-9]+[.,][0-9]{3,6})'
    ) 0.05 0.30

    # --- Potencia P1 (E/kW ano) 0-75: etiquetas P1 / Punta ---
    $r.p1 = Find-InRange $zonaPot @(
        '(?:^|[^A-Za-z])P1\b.*?([0-9]+[.,][0-9]+)',
        '[Pp]unta.*?([0-9]+[.,][0-9]+)',
        '[Pp]eriodo\s+1.*?([0-9]+[.,][0-9]+)'
    ) 0 75

    # --- Potencia P2 (E/kW ano) 0-75: etiquetas P2 / Valle / (a veces) P3 ---
    $r.p2 = Find-InRange $zonaPot @(
        '(?:^|[^A-Za-z])P2\b.*?([0-9]+[.,][0-9]+)',
        '[Vv]alle.*?([0-9]+[.,][0-9]+)',
        '(?:^|[^A-Za-z])P3\b.*?([0-9]+[.,][0-9]+)',
        '[Pp]eriodo\s+2.*?([0-9]+[.,][0-9]+)'
    ) 0 75

    if ($r.p1 -ne $null -and $r.p2 -eq $null) { $r.p2 = $r.p1 }
    return $r
}

function Get-WebText {
    param([string]$Url)
    try {
        $resp = Invoke-WebRequest -Uri $Url -Headers @{ 'User-Agent' = $UA } -TimeoutSec 15 -UseBasicParsing
        return [System.Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())
    } catch { return $null }
}

function Get-Holaluz {
    # Tarifa Clasica (precio unico 24h). Precios visibles en el HTML de /luz/.
    # Energia = primer valor EUR/kWh; potencia = primer valor EUR/kW (por dia) -> x365.
    $text = Get-WebText 'https://www.holaluz.com/luz/'
    if (-not $text) { return @{ energia=$null; p1=$null; p2=$null } }
    $t = $text -replace '<[^>]+>', ' '
    $e = Find-InRange $t @('([0-9][.,][0-9]{2,6})\s*[^0-9]{0,6}kWh') 0.05 0.30
    $pd = Find-InRange $t @('([0-9][.,][0-9]{2,6})\s*[^0-9]{0,6}kW(?!h|/kWh)') 0.01 0.60   # EUR/kW dia
    if ($e -ne $null -and $pd -ne $null) {
        $p = [Math]::Round($pd * 365, 2)
        return @{ energia=$e; p1=$p; p2=$p }
    }
    return @{ energia=$null; p1=$null; p2=$null }
}

function Get-Podo {
    # Tarifa "Precio Unico 24h": "Energia 24h X EUR kWh Potencia P1 Y EUR kW/dia P2 Z EUR kW/dia"
    $text = Get-WebText 'https://www.mipodo.com/tarifas-luz'
    if (-not $text) { return @{ energia=$null; p1=$null; p2=$null } }
    $t = ($text -replace '<[^>]+>', ' ') -replace '\s+', ' '
    $m = [regex]::Match($t, 'Energ.a\s*24h\s*([0-9][.,][0-9]{2,6})\s*[^0-9]{0,6}kWh\s*Potencia\s*P1\s*([0-9][.,][0-9]{2,6})\s*[^0-9]{0,10}kW.{0,4}d.a\s*P2\s*([0-9][.,][0-9]{2,6})', 'IgnoreCase')
    if ($m.Success) {
        $e  = Get-Number $m.Groups[1].Value
        $p1 = [Math]::Round((Get-Number $m.Groups[2].Value) * 365, 2)
        $p2 = [Math]::Round((Get-Number $m.Groups[3].Value) * 365, 2)
        if ($e -ge 0.05 -and $e -le 0.30) { return @{ energia=$e; p1=$p1; p2=$p2 } }
    }
    return @{ energia=$null; p1=$null; p2=$null }
}

function Get-Factor {
    $text = Get-WebText 'https://www.factorenergia.com/es/luz/tarifa-fija-de-luz-precio-unico/'
    if (-not $text) { return @{ energia=$null; p1=$null; p2=$null } }
    $mE = [regex]::Match($text, '(0[,\.]\d{3,4})\s*.{0,3}/kWh')
    $mP = [regex]::Match($text, '([\d,\.]+)\s*.{0,3}/kW\s*(?:d.a|day)', 'IgnoreCase')
    if ($mE.Success -and $mP.Success) {
        $e = Get-Number $mE.Groups[1].Value
        $p = (Get-Number $mP.Groups[1].Value) * 365
        return @{ energia=$e; p1=$p; p2=$p }
    }
    return @{ energia=$null; p1=$null; p2=$null }
}

function Get-Iberdrola {
    # Plan Estable (precio fijo). Energia 24h + Potencia Punta/Valle (EUR/kW dia -> x365).
    $text = Get-WebText 'https://www.iberdrola.es/luz/plan-estable'
    if (-not $text) { return @{ energia=$null; p1=$null; p2=$null } }
    $t = ($text -replace '<[^>]+>', ' ') -replace '\s+', ' '
    $mE  = [regex]::Match($t, '24\s*horas\s*del\s*d.a\s*([0-9][.,][0-9]{3,6})\s*[^0-9]{0,4}kWh', 'IgnoreCase')
    $mP1 = [regex]::Match($t, 'Periodo\s*Punta\s*([0-9][.,][0-9]{3,6})\s*[^0-9]{0,6}kW', 'IgnoreCase')
    $mP2 = [regex]::Match($t, 'Periodo\s*Valle\s*([0-9][.,][0-9]{3,6})\s*[^0-9]{0,6}kW', 'IgnoreCase')
    if ($mE.Success -and $mP1.Success -and $mP2.Success) {
        $e  = Get-Number $mE.Groups[1].Value
        $p1 = [Math]::Round((Get-Number $mP1.Groups[1].Value) * 365, 2)
        $p2 = [Math]::Round((Get-Number $mP2.Groups[1].Value) * 365, 2)
        if ($e -ge 0.05 -and $e -le 0.30) { return @{ energia=$e; p1=$p1; p2=$p2 } }
    }
    return @{ energia=$null; p1=$null; p2=$null }
}

# --------------------------------------------------------- EXCEL HELPERS -------

function Format-Sheet {
    <# Aplica cabecera con color, bordes y formato numerico a la hoja.
       $EnergyCols = columnas (1-based) de precios de energia -> 4 decimales; el resto 2. #>
    param($Ws, [int]$NumCols, [int]$LastRow, [int[]]$EnergyCols = @(2))

    $headerRange = $Ws.Range($Ws.Cells.Item(1,1), $Ws.Cells.Item(1,$NumCols))
    $headerRange.Interior.Color = 0x926036                 # BGR de 366092 (azul)
    $headerRange.Font.Color     = 0xFFFFFF
    $headerRange.Font.Bold      = $true
    $headerRange.HorizontalAlignment = -4108               # xlCenter

    if ($LastRow -ge 2) {
        # NumberFormatLocal usa el separador decimal local ($script:SepDec) para que
        # funcione tanto en Excel en espanol (',') como en ingles ('.').
        $sep  = $script:SepDec
        $fmt4 = "0${sep}0000"
        $fmt2 = "0${sep}00"
        # Col 1 = texto (Empresa). Cols de energia -> 4 decimales; resto -> 2 decimales.
        for ($c = 2; $c -le $NumCols; $c++) {
            $fmt = if ($EnergyCols -contains $c) { $fmt4 } else { $fmt2 }
            $Ws.Range($Ws.Cells.Item(2,$c), $Ws.Cells.Item($LastRow,$c)).NumberFormatLocal = $fmt
        }
        # Bordes en todo el rango usado
        $all = $Ws.Range($Ws.Cells.Item(1,1), $Ws.Cells.Item($LastRow,$NumCols))
        for ($b = 7; $b -le 12; $b++) { $all.Borders.Item($b).LineStyle = 1; $all.Borders.Item($b).Weight = 2 }
    }
    $Ws.Columns.Item(1).ColumnWidth = 22
    for ($c = 2; $c -le $NumCols; $c++) { $Ws.Columns.Item($c).ColumnWidth = 16 }
}

# --------------------------------------------------------- EXTRACCION ----------
# Log de cada ejecucion (util para tarea programada: ver que paso sin mirar la ventana)
$LogPath = Join-Path $PSScriptRoot 'benchmark_ultima_ejecucion.log'
try { Start-Transcript -Path $LogPath -Force | Out-Null } catch {}

Write-Host ("=" * 70)
Write-Host "CNMC BENCHMARK ENERGIA - DATOS REALES (API + WEB SCRAPING)"
Write-Host ("=" * 70)
Write-Host ""

$fechaHoy   = (Get-Date).ToString('yyyy-MM-dd')
$pestSemana = $fechaHoy   # nombre de pestana = fecha (formato YYYY-MM-DD, como el Excel de cowork)

Write-Host "Fecha: $fechaHoy"
Write-Host "Extrayendo $($TargetOffers.Count) empresas del API CNMC..."
Write-Host ""

$registros = New-Object System.Collections.ArrayList

try {
    $lista   = Invoke-CnmcApi "$BaseUrl/ofertas/electricidad?$(ConvertTo-QueryString (Get-FullParams 0))"
    # NO filtrar por precioUnico aqui: hay empresas presentes cuya mejor oferta no esta
    # marcada como precio unico (p.ej. Octopus, Nexus) y se perderian. Se filtra por empresa.
    $ofertas = @($lista.resultadoComparador)
    $nUnicas = @($ofertas | Where-Object { $_.tienePrecioUnico -eq 'S' }).Count
    Write-Host "  $($ofertas.Count) ofertas encontradas ($nUnicas de precio unico)"
    Write-Host ""
} catch {
    Write-Host "  ERROR al consultar el API: $($_.Exception.Message)"
    $ofertas = @()
}

# Emparejar cada marca con su oferta (nombre legal + cadena de busqueda)
foreach ($legal in $TargetOffers.Keys) {
    $cfg    = $TargetOffers[$legal]
    $marca  = $cfg.brand
    $search = $cfg.search

    $candidatas = @($ofertas | Where-Object { $_.comercializadora.ToUpper().Contains($legal.ToUpper()) })

    $registro = [ordered]@{ Empresa=$marca; Energia=$null; P1=$null; P2=$null; Fuente=$null }

    if ($candidatas.Count -gt 0) {
        # Prioridad: (1) precio unico + cadena de busqueda, (2) precio unico,
        # (3) cadena de busqueda, (4) primera candidata.
        $unicas = @($candidatas | Where-Object { $_.tienePrecioUnico -eq 'S' })
        $oferta = $unicas       | Where-Object { $_.oferta.ToLower().Contains($search.ToLower()) } | Select-Object -First 1
        if (-not $oferta) { $oferta = $unicas | Select-Object -First 1 }
        if (-not $oferta) { $oferta = $candidatas | Where-Object { $_.oferta.ToLower().Contains($search.ToLower()) } | Select-Object -First 1 }
        if (-not $oferta) { $oferta = $candidatas[0] }

        try {
            $det  = Invoke-CnmcApi "$BaseUrl/oferta?$(ConvertTo-QueryString (Get-FullParams $oferta.id))"
            $txt  = [string]$det.caracteristicas.caracteristicas
            $pr   = Parse-Prices $txt
            $registro.Energia = $pr.energia
            $registro.P1      = $pr.p1
            $registro.P2      = $pr.p2
            if ($pr.energia -ne $null) { $registro.Fuente = 'CNMC' }

            if ($pr.energia -ne $null -and $pr.p1 -ne $null -and $pr.p2 -ne $null) {
                Write-Host ("  OK   {0,-22} E:{1:N4} | P1:{2:N2} | P2:{3:N2}" -f $marca, $pr.energia, $pr.p1, $pr.p2)
            } else {
                Write-Host ("  WARN {0,-22} (parcial) E:{1} P1:{2} P2:{3}" -f $marca, $pr.energia, $pr.p1, $pr.p2)
            }
        } catch {
            Write-Host ("  ERR  {0,-22} {1}" -f $marca, $_.Exception.Message)
        }
    } else {
        Write-Host ("  --   {0,-22} (no encontrada en API)" -f $marca)
    }

    [void]$registros.Add($registro)
}

Write-Host ""
Write-Host "Web scraping (solo si el CNMC no las trajo): Holaluz, Podo, Factor, Iberdrola..."
Write-Host ""

# Estas ya son objetivos CNMC (arriba). Si el CNMC no las devolvio, se intenta su
# web y se RELLENA el registro existente (sin duplicar filas).
$scrapers = [ordered]@{
    'Holaluz'        = ${function:Get-Holaluz}
    'Podo'           = ${function:Get-Podo}
    'Factor Energia' = ${function:Get-Factor}
    'Iberdrola'      = ${function:Get-Iberdrola}
}
foreach ($emp in $scrapers.Keys) {
    $reg = $registros | Where-Object { $_.Empresa -eq $emp } | Select-Object -First 1

    if ($reg -and $reg.Energia -ne $null) {
        Write-Host ("  OK   {0,-22} (ya via CNMC)" -f $emp)
        continue
    }

    $pr = & $scrapers[$emp]
    if (-not $reg) {   # por si no estuviera como objetivo CNMC
        $reg = [ordered]@{ Empresa=$emp; Energia=$null; P1=$null; P2=$null; Fuente=$null }
        [void]$registros.Add($reg)
    }
    if ($pr.energia -ne $null) {
        $reg.Energia = $pr.energia; $reg.P1 = $pr.p1; $reg.P2 = $pr.p2; $reg.Fuente = 'Web'
    }

    if ($pr.energia -ne $null -and $pr.p1 -ne $null -and $pr.p2 -ne $null) {
        Write-Host ("  OK   {0,-22} E:{1:N4} | P1:{2:N2} | P2:{3:N2}" -f $emp, $pr.energia, $pr.p1, $pr.p2)
    } else {
        Write-Host ("  WARN {0,-22} (no extraido de web)" -f $emp)
    }
}

# ------------------------------------------------------------- EXCEL ----------
Write-Host ""
Write-Host "Actualizando Excel..."
Write-Host ""

if (-not (Test-Path $CarpetaSalida)) { New-Item -ItemType Directory -Path $CarpetaSalida -Force | Out-Null }
$rutaExcel = Join-Path $CarpetaSalida $NombreExcel

$xl = $null; $wb = $null
# PIDs de Excel ya abiertos (para no cerrar los libros del usuario despues)
$pidsPrevios = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$excelPid = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $script:SepDec = $xl.DecimalSeparator   # ',' en Excel espanol, '.' en ingles
    # Identificar el PID de ESTE Excel (el que no estaba abierto antes)
    Start-Sleep -Milliseconds 200
    $excelPid = @(Get-Process EXCEL -ErrorAction SilentlyContinue |
                  Where-Object { $_.Id -notin $pidsPrevios } |
                  Select-Object -ExpandProperty Id -First 1)

    $existia = Test-Path $rutaExcel
    $hojasDefecto = @()
    if ($existia) {
        $wb = $xl.Workbooks.Open($rutaExcel)
    } else {
        $wb = $xl.Workbooks.Add()
        $hojasDefecto = @($wb.Worksheets | ForEach-Object { $_.Name })   # p.ej. "Hoja1"
    }

    # --- Leer la semana anterior mas reciente (para arrastre y comparativa) ---
    $semanasPrev = @($wb.Worksheets | Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' -and $_.Name -ne $pestSemana } |
                     Sort-Object { $_.Name } -Descending)
    $datosAnt  = @{}
    $wsAntName = $null
    if ($semanasPrev.Count -gt 0) {
        $wsAnt = $semanasPrev[0]
        $wsAntName = $wsAnt.Name
        $lastRow = $wsAnt.UsedRange.Rows.Count
        for ($r = 2; $r -le $lastRow; $r++) {
            $nombre = $wsAnt.Cells.Item($r, 1).Value2
            if ($nombre) {
                $datosAnt[(Normalize-Name $nombre)] = @{   # clave normalizada (sin acentos)
                    energia = $wsAnt.Cells.Item($r, 2).Value2
                    p1      = $wsAnt.Cells.Item($r, 3).Value2
                    p2      = $wsAnt.Cells.Item($r, 4).Value2
                }
            }
        }
    }

    # --- Arrastre: empresas sin dato esta semana heredan el ultimo valor conocido ---
    if ($wsAntName) {
        foreach ($reg in $registros) {
            $clave = Normalize-Name $reg.Empresa
            if ($reg.Energia -eq $null -and $datosAnt.ContainsKey($clave)) {
                $ant = $datosAnt[$clave]
                if ($ant.energia -ne $null) {
                    $reg.Energia = $ant.energia
                    $reg.P1      = $ant.p1
                    $reg.P2      = $ant.p2
                    $reg.Fuente  = "Arrastrado ($wsAntName)"
                }
            }
        }
        $nArr = @($registros | Where-Object { "$($_.Fuente)" -like 'Arrastrado*' }).Count
        if ($nArr -gt 0) { Write-Host "  $nArr empresa(s) sin dato esta semana: arrastrado ultimo valor de '$wsAntName'" }
    }

    # --- Crear pestana semanal al principio (con nombre temporal) ---
    # Se crea ANTES de borrar la homonima previa para no dejar el libro sin hojas
    # (Excel prohibe libros sin al menos una hoja visible).
    $ws = $wb.Worksheets.Add($wb.Worksheets.Item(1))
    $ws.Name = "_tmp_$([Guid]::NewGuid().ToString('N').Substring(0,8))"

    # Borrar pestana semanal previa con el mismo nombre (idempotente)
    foreach ($sh in @($wb.Worksheets)) {
        if ($sh.Name -eq $pestSemana) { $sh.Delete() }
    }
    $ws.Name = $pestSemana

    $headers = @('Empresa', 'Precio energia (EUR/kWh)', 'P1 (EUR/kW/ano)', 'P2 (EUR/kW/ano)', 'Fuente')
    $nFilas  = $registros.Count + 1
    $arr = New-Object 'object[,]' $nFilas, 5
    for ($c = 0; $c -lt 5; $c++) { $arr[0, $c] = $headers[$c] }
    $i = 1
    foreach ($reg in $registros) {
        $arr[$i, 0] = $reg.Empresa
        $arr[$i, 1] = $reg.Energia
        $arr[$i, 2] = $reg.P1
        $arr[$i, 3] = $reg.P2
        $arr[$i, 4] = if ($reg.Fuente) { $reg.Fuente } else { 'No disponible' }
        $i++
    }
    $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($nFilas, 5)).Value2 = $arr

    # --- Formato pestana semanal (5 cols; la col 5 "Fuente" es texto) ---
    Format-Sheet $ws 5 $nFilas
    $ws.Columns.Item(5).ColumnWidth = 24

    Write-Host "  Pestana '$pestSemana' creada"

    # --- Comparativa vs semana anterior (reutiliza $datosAnt ya leido) ---
    if ($wsAntName) {
        Write-Host "  Creando comparativa vs '$wsAntName'..."

        $nombreComp = "$fechaHoy-Comparativa"
        foreach ($sh in @($wb.Worksheets)) { if ($sh.Name -eq $nombreComp) { $sh.Delete() } }
        $wsC = $wb.Worksheets.Add($wb.Worksheets.Item(2))
        $wsC.Name = $nombreComp

        $hc = @('Empresa','Energia Actual','Energia Anterior','Dif (EUR/kWh)','Cambio %',
                'P1 Actual','P1 Anterior','Dif (EUR/kW/ano)','Cambio %',
                'P2 Actual','P2 Anterior','Dif (EUR/kW/ano)','Cambio %')

        # TODAS las empresas (aunque falten en una de las semanas -> anterior en blanco)
        $filas = @($registros)
        $nC = $filas.Count + 1
        $arrC = New-Object 'object[,]' $nC, 13
        for ($c = 0; $c -lt 13; $c++) { $arrC[0, $c] = $hc[$c] }

        $i = 1
        foreach ($reg in $filas) {
            $k = Normalize-Name $reg.Empresa
            $ant = if ($datosAnt.ContainsKey($k)) { $datosAnt[$k] } else { @{ energia=$null; p1=$null; p2=$null } }
            $arrC[$i, 0] = $reg.Empresa
            # bloques: (col base 0/4/8) Actual, Anterior, Dif, %Cambio
            $bloques = @(
                @{ base = 1;  act = $reg.Energia; prev = $ant.energia },
                @{ base = 5;  act = $reg.P1;      prev = $ant.p1 },
                @{ base = 9;  act = $reg.P2;      prev = $ant.p2 }
            )
            foreach ($b in $bloques) {
                $base = [int]$b.base
                $arrC[$i, $base]       = $b.act
                $arrC[$i, ($base + 1)] = $b.prev
                if ($b.act -ne $null -and $b.prev -ne $null -and $b.prev -ne 0) {
                    $dif = $b.act - $b.prev
                    $arrC[$i, ($base + 2)] = $dif
                    $arrC[$i, ($base + 3)] = ($dif / $b.prev * 100)
                }
            }
            $i++
        }
        $wsC.Range($wsC.Cells.Item(1, 1), $wsC.Cells.Item($nC, 13)).Value2 = $arrC

        Format-Sheet $wsC 13 $nC @(2, 3, 4)   # cols de energia (Actual/Anterior/Dif) a 4 decimales
        Write-Host "  Comparativa creada"
    }

    # Eliminar hojas por defecto ("Hoja1") de un libro recien creado
    if ($hojasDefecto.Count -gt 0) {
        foreach ($sh in @($wb.Worksheets)) { if ($hojasDefecto -contains $sh.Name) { $sh.Delete() } }
    }

    if ($existia) { $wb.Save() } else { $wb.SaveAs($rutaExcel, 51) }  # 51 = xlsx
    Write-Host ""
    Write-Host "  Pestanas: $(@($wb.Worksheets | ForEach-Object { $_.Name }) -join ', ')"
}
finally {
    foreach ($o in @($ws, $wsC)) {
        if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
    }
    if ($wb) { try { $wb.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {} }
    if ($xl) { try { $xl.Quit();        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    # Garantia: si nuestro Excel sigue vivo, cerrar SOLO ese proceso (por PID)
    if ($excelPid) {
        Get-Process -Id $excelPid -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ""
Write-Host "Archivo guardado en: $rutaExcel"
Write-Host "Total empresas: $($registros.Count)"
try { Stop-Transcript | Out-Null } catch {}
