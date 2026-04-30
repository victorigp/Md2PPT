param([string]$DocsDir)

# Obtener el primer .md por fecha (mas antiguo)
$md = Get-ChildItem -Path $DocsDir -Filter "*.md" | Sort-Object LastWriteTime | Select-Object -First 1
if (-not $md) {
    Write-Error "No se encontro ningun archivo .md en $DocsDir"
    exit 1
}

# Extraer titulo H1 del markdown primero para poder excluir el fichero de salida
$linea = Get-Content $md.FullName -Encoding UTF8 | Where-Object { $_ -match '^# [^#]' } | Select-Object -First 1
if ($linea) {
    $nombre = $linea -replace '^# ', ''
    $nombre = $nombre -replace '[\\/:*?"<>|]', ''
    $nombre = $nombre.Trim()
    if (-not $nombre) { $nombre = "resultado" }
} else {
    $nombre = "resultado"
}

# Obtener el primer .pptx por fecha (mas antiguo), excluyendo ficheros generados previamente
$pptx = Get-ChildItem -Path $DocsDir -Filter "*.pptx" | Where-Object {
    $_.BaseName -ne $nombre -and $_.BaseName -notmatch "^${nombre}_\d+$"
} | Sort-Object LastWriteTime | Select-Object -First 1
if (-not $pptx) {
    Write-Error "No se encontro ningun archivo .pptx (plantilla) en $DocsDir"
    exit 1
}

# Comprobar si ya existe el fichero de salida y anadir sufijo incremental
$salida = Join-Path $DocsDir "$nombre.pptx"
if (Test-Path $salida) {
    $i = 1
    while (Test-Path (Join-Path $DocsDir "${nombre}_$i.pptx")) {
        $i++
    }
    $nombre = "${nombre}_$i"
}

# Devolver: md|pptx|salida
Write-Output "$($md.Name)|$($pptx.Name)|$nombre"
