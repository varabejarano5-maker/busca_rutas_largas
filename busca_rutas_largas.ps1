<#
.SYNOPSIS
    Buscador de rutas de archivos que exceden el límite de caracteres en Windows.

.DESCRIPTION
    Script con interfaz gráfica para auditar directorios, identificar rutas extensas y abrir su ubicación en el Explorador de archivos.

.NOTES
    Versión: 1.0.0
    Fecha:   2026-10-10
    Autor:   varabejarano5-maker
#>


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Funcion de Windows que usa el propio Explorador en "Abrir ubicacion del archivo":
# abre la carpeta y deja el archivo seleccionado (tambien si la ventana ya estaba abierta).
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class ExploradorSel {
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    static extern IntPtr ILCreateFromPathW(string pszPath);
    [DllImport("shell32.dll")]
    static extern void ILFree(IntPtr pidl);
    [DllImport("shell32.dll")]
    static extern int SHOpenFolderAndSelectItems(IntPtr pidlFolder, uint cidl, IntPtr apidl, uint dwFlags);
    public static bool Seleccionar(string ruta) {
        IntPtr pidl = ILCreateFromPathW(ruta);
        if (pidl == IntPtr.Zero) return false;
        try { return SHOpenFolderAndSelectItems(pidl, 0, IntPtr.Zero, 0) == 0; }
        finally { ILFree(pidl); }
    }
}
"@

# ---------------------------------------------------------------------------
# Seleccion de carpeta
# ---------------------------------------------------------------------------
function Seleccionar-Carpeta {
    $dialogo = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialogo.Description = "Selecciona la carpeta principal a analizar"
    if ($dialogo.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialogo.SelectedPath
    }
    return $null
}

$script:carpeta = Seleccionar-Carpeta
if (-not $script:carpeta) {
    Write-Host "Operación cancelada." -ForegroundColor Yellow
    exit
}

# ---------------------------------------------------------------------------
# Tabla de datos (mismas columnas que antes)
# ---------------------------------------------------------------------------
$tabla = New-Object System.Data.DataTable
[void]$tabla.Columns.Add("Caracteres",    [int])
[void]$tabla.Columns.Add("Extension",     [string])
[void]$tabla.Columns.Add("NombreArchivo", [string])
[void]$tabla.Columns.Add("RutaCompleta",  [string])
$tabla.DefaultView.Sort = "Caracteres DESC"

# ---------------------------------------------------------------------------
# Ventana
# ---------------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(1200, 700)
$form.StartPosition = "CenterScreen"
$form.KeyPreview = $true

# Panel superior: filtro
$panelSup = New-Object System.Windows.Forms.Panel
$panelSup.Dock = "Top"
$panelSup.Height = 38

$lblFiltro = New-Object System.Windows.Forms.Label
$lblFiltro.Text = "Filtrar:"
$lblFiltro.AutoSize = $true
$lblFiltro.Location = New-Object System.Drawing.Point(10, 11)

$txtFiltro = New-Object System.Windows.Forms.TextBox
$txtFiltro.Location = New-Object System.Drawing.Point(60, 8)
$txtFiltro.Width = 400

$lblAyuda = New-Object System.Windows.Forms.Label
$lblAyuda.Text = "Doble clic o Enter = abrir ubicación   |   F5 = actualizar   |   Clic en el encabezado = ordenar"
$lblAyuda.AutoSize = $true
$lblAyuda.ForeColor = [System.Drawing.Color]::DimGray
$lblAyuda.Location = New-Object System.Drawing.Point(475, 11)

$panelSup.Controls.AddRange(@($lblFiltro, $txtFiltro, $lblAyuda))

# Tabla
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = "Fill"
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.RowHeadersVisible = $true
$grid.RowHeadersWidthSizeMode = "DisableResizing"
$grid.RowHeadersWidth = 60
$grid.TopLeftHeaderCell.Value = "N°"
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $true
$grid.BackgroundColor = [System.Drawing.Color]::White

# Panel inferior: botones + estado
$panelInf = New-Object System.Windows.Forms.FlowLayoutPanel
$panelInf.Dock = "Bottom"
$panelInf.Height = 44
$panelInf.Padding = New-Object System.Windows.Forms.Padding(6)

function Nuevo-Boton($texto, $ancho) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $texto
    $b.Width = $ancho
    $b.Height = 30
    return $b
}
$btnAbrir      = Nuevo-Boton "Abrir ubicación"   130
$btnActualizar = Nuevo-Boton "Actualizar (F5)"   130
$btnCarpeta    = Nuevo-Boton "Cambiar carpeta..." 130
$btnSalir      = Nuevo-Boton "Salir"             90

$lblEstado = New-Object System.Windows.Forms.Label
$lblEstado.AutoSize = $true
$lblEstado.Margin = New-Object System.Windows.Forms.Padding(15, 9, 0, 0)

$panelInf.Controls.AddRange(@($btnAbrir, $btnActualizar, $btnCarpeta, $btnSalir, $lblEstado))

$form.Controls.Add($grid)
$form.Controls.Add($panelSup)
$form.Controls.Add($panelInf)
$grid.BringToFront()

# ---------------------------------------------------------------------------
# Funciones
# ---------------------------------------------------------------------------
function Escapar-Like([string]$t) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $t.ToCharArray()) {
        switch ($c) {
            "'"     { [void]$sb.Append("''") }
            '['     { [void]$sb.Append('[[]') }
            ']'     { [void]$sb.Append('[]]') }
            '*'     { [void]$sb.Append('[*]') }
            '%'     { [void]$sb.Append('[%]') }
            default { [void]$sb.Append($c) }
        }
    }
    return $sb.ToString()
}

function Aplicar-Filtro {
    $t = $txtFiltro.Text.Trim()
    if ($t -eq "") {
        $tabla.DefaultView.RowFilter = ""
    } else {
        $e = Escapar-Like $t
        $tabla.DefaultView.RowFilter = "NombreArchivo LIKE '%$e%' OR RutaCompleta LIKE '%$e%' OR Extension LIKE '%$e%'"
    }
    Actualizar-Estado
}

function Actualizar-Estado {
    $total   = $tabla.Rows.Count
    $visibles = $tabla.DefaultView.Count
    $largos  = @($tabla.Select("Caracteres > 260")).Count
    $txt = "$total archivos   |   $largos con ruta mayor a 260 caracteres   |   Actualizado: $($script:horaCarga)"
    if ($visibles -ne $total) { $txt = "Mostrando $visibles de " + $txt }
    $lblEstado.Text = $txt
}

function Cargar-Archivos {
    $form.Text = "Archivos ordenados por longitud de ruta - $($script:carpeta)"
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $lblEstado.Text = "Analizando, espera un momento..."
    $form.Refresh()

    # Recordar posicion de la lista para no perderla al actualizar
    $posScroll = $grid.FirstDisplayedScrollingRowIndex

    $archivos = Get-ChildItem -LiteralPath $script:carpeta -Recurse -File -Force -ErrorAction SilentlyContinue

    $grid.DataSource = $null
    $tabla.BeginLoadData()
    $tabla.Clear()
    foreach ($f in $archivos) {
        [void]$tabla.Rows.Add($f.FullName.Length, $f.Extension.ToLower(), $f.Name, $f.FullName)
    }
    $tabla.EndLoadData()
    $grid.DataSource = $tabla.DefaultView

    $grid.Columns["Caracteres"].Width    = 80
    $grid.Columns["Extension"].Width     = 80
    $grid.Columns["NombreArchivo"].Width = 280
    $grid.Columns["RutaCompleta"].AutoSizeMode = "Fill"

    # Ancho de la columna de numeracion segun la cantidad de archivos
    $anchoNum = [System.Windows.Forms.TextRenderer]::MeasureText(($tabla.Rows.Count + 1).ToString(), $grid.Font).Width
    $grid.RowHeadersWidth = [Math]::Max(60, $anchoNum + 30)

    if ($posScroll -ge 0 -and $grid.Rows.Count -gt 0) {
        try { $grid.FirstDisplayedScrollingRowIndex = [Math]::Min($posScroll, $grid.Rows.Count - 1) } catch {}
    }
    $grid.ClearSelection()

    $script:horaCarga = (Get-Date).ToString("HH:mm:ss")
    Actualizar-Estado
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
}

function Ruta-Larga([string]$ruta) {
    # Prefijo que permite a Windows manejar rutas de mas de 260 caracteres
    if ($ruta.StartsWith("\\?\")) { return $ruta }
    if ($ruta.StartsWith("\\"))     { return "\\?\UNC\" + $ruta.Substring(2) }
    return "\\?\" + $ruta
}

function Existe-Archivo([string]$ruta) {
    if (Test-Path -LiteralPath $ruta -PathType Leaf) { return $true }
    try { return [System.IO.File]::Exists((Ruta-Larga $ruta)) } catch { return $false }
}

function Seleccionar-ConShell([string]$ruta) {
    # Respaldo: abre la carpeta y selecciona el archivo desde la propia ventana del Explorador
    $i = $ruta.LastIndexOf("\")
    $carpetaArch = $ruta.Substring(0, $i)
    $nombre      = $ruta.Substring($i + 1)
    try {
        $shell = New-Object -ComObject Shell.Application
        Start-Process explorer.exe -ArgumentList "`"$carpetaArch`""
        for ($n = 0; $n -lt 40; $n++) {
            Start-Sleep -Milliseconds 150
            foreach ($w in @($shell.Windows())) {
                try {
                    if ($w.Document.Folder.Self.Path -eq $carpetaArch) {
                        $item = $w.Document.Folder.ParseName($nombre)
                        if ($item) {
                            # 1 seleccionar + 4 quitar otras selecciones + 8 hacer visible + 16 enfocar
                            $w.Document.SelectItem($item, 29)
                            return $true
                        }
                    }
                } catch {}
            }
        }
    } catch {}
    return $false
}

function Mostrar-EnExplorador([string]$ruta) {
    if ([ExploradorSel]::Seleccionar($ruta))              { return }
    if ([ExploradorSel]::Seleccionar((Ruta-Larga $ruta))) { return }
    if (Seleccionar-ConShell $ruta)                       { return }
    Start-Process explorer.exe -ArgumentList "/select,`"$ruta`""
}

function Abrir-Ubicacion {
    $filas = @($grid.SelectedRows)
    if ($filas.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("Selecciona al menos un archivo de la lista.", "Abrir ubicación")
        return
    }
    if ($filas.Count -gt 10) {
        $resp = [System.Windows.Forms.MessageBox]::Show(
            "Seleccionaste $($filas.Count) archivos. Se abrirá una ventana del Explorador por cada uno. ¿Continuar?",
            "Confirmar", [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($resp -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    $noEncontrados = @()
    foreach ($fila in $filas) {
        $ruta = [string]$fila.Cells["RutaCompleta"].Value
        if (Existe-Archivo $ruta) {
            Mostrar-EnExplorador $ruta
        } else {
            $noEncontrados += $ruta
        }
    }
    if ($noEncontrados.Count -gt 0) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Estos archivos ya no existen (¿fueron movidos, renombrados o eliminados?).`nPulsa 'Actualizar' para refrescar la lista.`n`n" + ($noEncontrados -join "`n"),
            "Archivo no encontrado", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
}

# ---------------------------------------------------------------------------
# Eventos
# ---------------------------------------------------------------------------
$btnAbrir.Add_Click({ Abrir-Ubicacion })
$btnActualizar.Add_Click({ Cargar-Archivos })
$btnCarpeta.Add_Click({
    $nueva = Seleccionar-Carpeta
    if ($nueva) {
        $script:carpeta = $nueva
        $grid.DataSource = $null
        Cargar-Archivos
    }
})
$btnSalir.Add_Click({ $form.Close() })

$txtFiltro.Add_TextChanged({ Aplicar-Filtro })

$grid.Add_CellDoubleClick({
    param($s, $e)
    if ($e.RowIndex -ge 0) { Abrir-Ubicacion }
})
$grid.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        Abrir-Ubicacion
    }
})
$form.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::F5) {
        $e.SuppressKeyPress = $true
        Cargar-Archivos
    }
})

# Numeracion en la cabecera de cada fila (sigue el orden que se ve en pantalla,
# tambien despues de ordenar por otra columna o de filtrar)
$grid.Add_RowPostPaint({
    param($s, $e)
    $rect = New-Object System.Drawing.Rectangle($e.RowBounds.Left, $e.RowBounds.Top, ($s.RowHeadersWidth - 6), $e.RowBounds.Height)
    $flags = [System.Windows.Forms.TextFormatFlags]::Right -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter
    [System.Windows.Forms.TextRenderer]::DrawText($e.Graphics, ($e.RowIndex + 1).ToString(), $s.Font, $rect, [System.Drawing.Color]::DimGray, $flags)
})

# Cargar la lista apenas se muestre la ventana
$form.Add_Shown({ Cargar-Archivos })

[void]$form.ShowDialog()
$form.Dispose()
