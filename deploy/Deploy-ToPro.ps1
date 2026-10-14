<#
.SYNOPSIS
    Copia el contenido de app/ (ToperBTR) a la ruta fisica del sitio IIS de PRODUCCION,
    excluyendo lo que no debe desplegarse ahi (ver docs/14-despliegue-produccion.md).

.DESCRIPTION
    Excluye a proposito:
      - conn_test.asp       -> diagnostico de BD sin ninguna autenticacion, no debe existir en PRO
      - externos\           -> endpoint de intercambio con operadores, fuera de esta ola de despliegue
      - global.asa          -> el de este repo/entorno tiene las credenciales de eir_test; PRO necesita
                               uno propio, construido a mano con las credenciales reales (ver docs/14)
      - global.asa.example  -> es solo la plantilla de referencia del repo, no hace falta en el servidor

    Usa robocopy /E (copia todo, incluidas subcarpetas vacias) SIN /MIR: solo añade y actualiza
    ficheros, nunca borra nada que ya exista en el destino. Es deliberado - no queremos arriesgarnos
    a que una sincronizacion "espejo" borre el global.asa de PRO (que vive fuera de este repo) u
    otro fichero propio del servidor que no este en este origen.

    Por defecto se ejecuta en modo SIMULACION (robocopy /L): no copia nada, solo dice que haria.
    Hay que pasar -Ejecutar explicitamente para copiar de verdad.

.PARAMETER Destino
    Ruta fisica del sitio IIS de PRO (UNC o local), p.ej.:
      .\Deploy-ToPro.ps1 -Destino "\\servidor-pro\sitios\ToperBTR"

.PARAMETER Origen
    Carpeta app/ de origen. Por defecto, la que esta junto a este script (..\app).

.PARAMETER Ejecutar
    Sin este switch, el script SOLO SIMULA (no copia nada). Pasarlo para copiar de verdad.

.EXAMPLE
    # Ver que haria, sin copiar nada todavia:
    .\Deploy-ToPro.ps1 -Destino "\\servidor-pro\sitios\ToperBTR"

.EXAMPLE
    # Copiar de verdad, una vez revisada la simulacion:
    .\Deploy-ToPro.ps1 -Destino "\\servidor-pro\sitios\ToperBTR" -Ejecutar
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$Destino,

    [string]$Origen = (Join-Path $PSScriptRoot "..\app"),

    [switch]$Ejecutar
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $Origen)) {
    Write-Error "No se encuentra el origen '$Origen'. Pasa -Origen con la ruta correcta a app/."
    exit 1
}
$Origen = (Resolve-Path $Origen).Path

if (-not (Test-Path $Destino)) {
    Write-Warning "El destino '$Destino' no existe o no es accesible desde esta maquina."
    Write-Warning "Revisa la ruta/permisos antes de continuar (hace falta estar en la VPN, o mapear la unidad primero?)."
    exit 1
}

$exclFiles = @("conn_test.asp", "global.asa", "global.asa.example", "Thumbs.db", "desktop.ini")
$exclDirs  = @("externos", ".git")

Write-Host "================================================================"
Write-Host " Despliegue ToperBTR -> PRO"
Write-Host "================================================================"
Write-Host "Origen:  $Origen"
Write-Host "Destino: $Destino"
Write-Host "Ficheros excluidos: $($exclFiles -join ', ')"
Write-Host "Carpetas excluidas: $($exclDirs -join ', ')"
Write-Host ""

$robocopyArgs = @($Origen, $Destino, "/E")   # /E: subcarpetas incluidas, incluso vacias - sin /MIR, no borra nada
$robocopyArgs += "/XF"
$robocopyArgs += $exclFiles
$robocopyArgs += "/XD"
$robocopyArgs += $exclDirs
$robocopyArgs += @("/NFL", "/NDL", "/NP")    # salida resumida: sin listar cada fichero/carpeta ni barra de progreso

if (-not $Ejecutar) {
    $robocopyArgs += "/L"
    Write-Host "*** MODO SIMULACION *** No se va a copiar nada todavia." -ForegroundColor Yellow
    Write-Host "Revisa el resultado y vuelve a lanzar con -Ejecutar para copiar de verdad." -ForegroundColor Yellow
    Write-Host ""
}

& robocopy @robocopyArgs
$codigo = $LASTEXITCODE

Write-Host ""
if ($codigo -ge 8) {
    Write-Warning "Robocopy devolvio el codigo $codigo (>= 8 = hubo errores). Revisa la salida de arriba."
} else {
    Write-Host "Robocopy termino sin errores (codigo $codigo)." -ForegroundColor Green
}

Write-Host ""
Write-Host "Este script NO hace lo siguiente (ver docs/14-despliegue-produccion.md):" -ForegroundColor Cyan
Write-Host " - No toca global.asa en destino: construyelo/confirmalo a mano con las credenciales reales de PRO."
Write-Host " - No ejecuta nada de db/ contra ninguna base de datos."
Write-Host " - No valida el modo de pipeline de IIS ni la cabecera HSTS - revisalo aparte."
