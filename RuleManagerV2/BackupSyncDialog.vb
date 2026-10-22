Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Text
Imports System.Threading.Tasks

' Alias explícitos a nivel de archivo: tienen máxima precedencia en el compilador VB.NET
' y eliminan de raíz cualquier colisión con la API de Autodesk Inventor (Color, Point, TextBox, ProgressBar, Path, File, Environment)
Imports Color = System.Drawing.Color
Imports Point = System.Drawing.Point
Imports Font = System.Drawing.Font
Imports Size = System.Drawing.Size
Imports FontStyle = System.Drawing.FontStyle
Imports TextBox = System.Windows.Forms.TextBox
Imports ProgressBar = System.Windows.Forms.ProgressBar
Imports Button = System.Windows.Forms.Button
Imports Label = System.Windows.Forms.Label
Imports ComboBox = System.Windows.Forms.ComboBox
Imports CheckBox = System.Windows.Forms.CheckBox
Imports TreeView = System.Windows.Forms.TreeView
Imports TreeNode = System.Windows.Forms.TreeNode
Imports TreeNodeCollection = System.Windows.Forms.TreeNodeCollection
Imports TreeViewEventArgs = System.Windows.Forms.TreeViewEventArgs
Imports Panel = System.Windows.Forms.Panel
Imports Splitter = System.Windows.Forms.Splitter
Imports Form = System.Windows.Forms.Form
Imports FolderBrowserDialog = System.Windows.Forms.FolderBrowserDialog
Imports MessageBox = System.Windows.Forms.MessageBox
Imports MessageBoxButtons = System.Windows.Forms.MessageBoxButtons
Imports MessageBoxIcon = System.Windows.Forms.MessageBoxIcon
Imports DialogResult = System.Windows.Forms.DialogResult
Imports Cursors = System.Windows.Forms.Cursors
Imports FlatStyle = System.Windows.Forms.FlatStyle
Imports FormStartPosition = System.Windows.Forms.FormStartPosition
Imports ProgressBarStyle = System.Windows.Forms.ProgressBarStyle
Imports Orientation = System.Windows.Forms.Orientation
Imports DockStyle = System.Windows.Forms.DockStyle
Imports Padding = System.Windows.Forms.Padding
Imports ScrollBars = System.Windows.Forms.ScrollBars
Imports BorderStyle = System.Windows.Forms.BorderStyle
Imports ComboBoxStyle = System.Windows.Forms.ComboBoxStyle
Imports Path = System.IO.Path
Imports File = System.IO.File
Imports Directory = System.IO.Directory
Imports SysEnv = System.Environment

''' <summary>
''' Diálogo interactivo para gestión de Backup y Sincronización de Bibliotecas (Local vs Servidor de Red).
''' Visualización jerárquica tipo árbol (TreeView) optimizada:
''' - Ramas 100% idénticas se muestran agrupadas y colapsadas para evitar ruido visual.
''' - Ramas con diferencias se auto-despliegan y posicionan el foco automáticamente en los archivos a sincronizar.
''' - Motor multihilo para saturar redes 5GbE / Gigabit sin congelar la interfaz.
''' </summary>
Public Class BackupSyncDialog
    Inherits Form

    Private ReadOnly _invApp As Inventor.Application
    Private ReadOnly _syncManager As LibraryBackupSyncManager

    Private _txtRutaLocal As TextBox
    Private _txtRutaServidor As TextBox
    Private _cmbIntervalo As ComboBox
    Private _chkAutoSync As CheckBox
    Private _btnExaminarLocal As Button
    Private _btnExaminarServidor As Button
    Private _btnComparar As Button
    Private _btnPull As Button
    Private _btnPush As Button
    Private _btnAbrirLogs As Button
    Private _btnCerrar As Button

    ' Controles del Árbol y Filtrado
    Private _treeView As TreeView
    Private _chkSoloDiferencias As CheckBox
    Private _btnColapsarIdenticas As Button
    Private _btnExpandirDiferencias As Button

    ' Panel de Detalles
    Private _lblDetalleTitulo As Label
    Private _lblDetalleEstado As Label
    Private _lblDetalleRutas As Label
    Private _lblDetalleMetadatos As Label
    Private _txtDetallesInternos As TextBox
    Private _lblResumen As Label
    Private _lblEstadoRed As Label
    Private _progressBar As ProgressBar

    Private _itemsComparados As New List(Of FileComparisonItem)
    Private _estaOcupado As Boolean = False

    Public Sub New(Optional invApp As Inventor.Application = Nothing)
        If invApp IsNot Nothing Then
            _invApp = invApp
        Else
            Try
                _invApp = DirectCast(Microsoft.VisualBasic.Interaction.GetObject(Nothing, "Inventor.Application"), Inventor.Application)
            Catch
                _invApp = Nothing
            End Try
        End If
        _syncManager = New LibraryBackupSyncManager(_invApp)
        InitializeComponent()
        CargarConfiguracion()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "RuleManager V2 - Sincronización y Backup de Bibliotecas CAD (Árbol Jerárquico 5GbE)"
        Me.Size = New Size(1160, 760)
        Me.MinimumSize = New Size(960, 640)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Font = New Font("Segoe UI", 9.0F)
        Me.BackColor = Color.FromArgb(248, 249, 250)

        ' Header Panel
        Dim pnlHeader As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 72,
            .BackColor = Color.FromArgb(15, 23, 42),
            .Padding = New Padding(20, 12, 20, 12)
        }

        Dim lblTitle As New Label With {
            .Text = "Backup & Sincronización de Bibliotecas CAD (Local ⇄ Servidor)",
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 11.5F, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(20, 12)
        }

        Dim lblSubtitle As New Label With {
            .Text = "Vista en Árbol por Ramas: Las carpetas idénticas permanecen colapsadas. Las carpetas con diferencias se auto-despliegan al instante.",
            .ForeColor = Color.FromArgb(148, 163, 184),
            .Font = New Font("Segoe UI", 8.5F),
            .AutoSize = True,
            .Location = New Point(20, 38)
        }

        _lblEstadoRed = New Label With {
            .Text = "● Comprobando red...",
            .ForeColor = Color.Gold,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .AutoSize = True,
            .Dock = DockStyle.Right
        }

        pnlHeader.Controls.AddRange({lblTitle, lblSubtitle, _lblEstadoRed})

        ' Top Config Panel
        Dim pnlConfig As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 114,
            .BackColor = Color.White,
            .Padding = New Padding(20, 10, 20, 10)
        }

        ' Fila 1: Ruta Local
        Dim lblLocal As New Label With {.Text = "Carpeta Local:", .Location = New Point(20, 14), .AutoSize = True, .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)}
        _txtRutaLocal = New TextBox With {.Location = New Point(130, 11), .Width = 450}
        _btnExaminarLocal = New Button With {.Text = "Examinar...", .Location = New Point(590, 10), .Width = 90, .BackColor = Color.White, .FlatStyle = FlatStyle.Flat}
        AddHandler _btnExaminarLocal.Click, AddressOf OnExaminarLocalClicked

        ' Fila 2: Ruta Servidor
        Dim lblServidor As New Label With {.Text = "Ruta Servidor:", .Location = New Point(20, 44), .AutoSize = True, .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)}
        _txtRutaServidor = New TextBox With {.Location = New Point(130, 41), .Width = 450}
        _btnExaminarServidor = New Button With {.Text = "Examinar...", .Location = New Point(590, 40), .Width = 90, .BackColor = Color.White, .FlatStyle = FlatStyle.Flat}
        AddHandler _btnExaminarServidor.Click, AddressOf OnExaminarServidorClicked

        ' Configuración de intervalo y arranque
        Dim lblInt As New Label With {.Text = "Frecuencia:", .Location = New Point(700, 14), .AutoSize = True}
        _cmbIntervalo = New ComboBox With {.Location = New Point(780, 11), .Width = 110, .DropDownStyle = ComboBoxStyle.DropDownList}
        _cmbIntervalo.Items.AddRange({"Diario", "Semanal", "Mensual"})
        _cmbIntervalo.SelectedIndex = 0
        AddHandler _cmbIntervalo.SelectedIndexChanged, Sub() GuardarConfiguracion()

        _chkAutoSync = New CheckBox With {
            .Text = "Sincronizar automáticamente al iniciar Inventor",
            .Location = New Point(700, 42),
            .AutoSize = True
        }
        AddHandler _chkAutoSync.CheckedChanged, Sub() GuardarConfiguracion()

        ' Fila de Acciones
        _btnComparar = New Button With {
            .Text = "⚡ Comprobar Diferencias (Alta Velocidad)",
            .Location = New Point(20, 74),
            .Size = New Size(260, 30),
            .BackColor = Color.FromArgb(224, 242, 254),
            .ForeColor = Color.FromArgb(3, 105, 161),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .Cursor = Cursors.Hand
        }
        _btnComparar.FlatAppearance.BorderColor = Color.FromArgb(125, 211, 252)
        AddHandler _btnComparar.Click, AddressOf OnCompararClicked

        _btnPull = New Button With {
            .Text = "⬇ Hacer PULL (Servidor -> Local)",
            .Location = New Point(290, 74),
            .Size = New Size(220, 30),
            .BackColor = Color.FromArgb(238, 242, 255),
            .ForeColor = Color.FromArgb(67, 56, 202),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .Cursor = Cursors.Hand,
            .Enabled = False
        }
        _btnPull.FlatAppearance.BorderColor = Color.FromArgb(165, 180, 252)
        AddHandler _btnPull.Click, AddressOf OnPullClicked

        _btnPush = New Button With {
            .Text = "⬆ Hacer PUSH (Local -> Servidor)",
            .Location = New Point(520, 74),
            .Size = New Size(220, 30),
            .BackColor = Color.FromArgb(254, 243, 199),
            .ForeColor = Color.FromArgb(146, 64, 14),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .Cursor = Cursors.Hand,
            .Enabled = False
        }
        _btnPush.FlatAppearance.BorderColor = Color.FromArgb(252, 211, 77)
        AddHandler _btnPush.Click, AddressOf OnPushClicked

        _btnAbrirLogs = New Button With {
            .Text = "📁 Ver Logs en Servidor",
            .Location = New Point(750, 74),
            .Size = New Size(160, 30),
            .BackColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        AddHandler _btnAbrirLogs.Click, AddressOf OnAbrirLogsClicked

        pnlConfig.Controls.AddRange({
            lblLocal, _txtRutaLocal, _btnExaminarLocal,
            lblServidor, _txtRutaServidor, _btnExaminarServidor,
            lblInt, _cmbIntervalo, _chkAutoSync,
            _btnComparar, _btnPull, _btnPush, _btnAbrirLogs
        })

        ' ==============================================================
        ' CONTENEDOR CENTRAL: Árbol (Izquierda) + Divisor Splitter + Detalles (Derecha)
        ' Utiliza Splitter nativo estándar de Windows Forms (100% estable, sin SplitterDistance)
        ' ==============================================================
        Dim pnlCentral As New Panel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.FromArgb(248, 250, 252)
        }

        ' PANEL IZQUIERDO: Árbol de Carpetas/Archivos + Acciones de Navegación
        Dim pnlTreeHost As New Panel With {
            .Dock = DockStyle.Left,
            .Width = 500,
            .MinimumSize = New Size(430, 0),
            .Padding = New Padding(20, 10, 10, 5)
        }

        Dim pnlTreeToolbar As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 36,
            .BackColor = Color.FromArgb(241, 245, 249),
            .Padding = New Padding(6, 4, 6, 4)
        }

        _chkSoloDiferencias = New CheckBox With {
            .Text = "Solo ramas con diferencias",
            .Dock = DockStyle.Left,
            .AutoSize = True,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(30, 41, 59),
            .Checked = False
        }
        AddHandler _chkSoloDiferencias.CheckedChanged, Sub() AplicarFiltroArbol()

        ' Botón Colapsar Idénticas (se añade primero para quedar a la derecha)
        _btnColapsarIdenticas = New Button With {
            .Text = "Colapsar idénticas",
            .Dock = DockStyle.Right,
            .Width = 125,
            .BackColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 8.25F),
            .Cursor = Cursors.Hand
        }
        _btnColapsarIdenticas.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        AddHandler _btnColapsarIdenticas.Click, Sub() ColapsarRamasIdenticas()

        ' Botón Expandir Diferencias (se añade después para quedar a la izquierda de Colapsar)
        _btnExpandirDiferencias = New Button With {
            .Text = "Expandir diferencias",
            .Dock = DockStyle.Right,
            .Width = 140,
            .BackColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 8.25F),
            .Cursor = Cursors.Hand
        }
        _btnExpandirDiferencias.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        AddHandler _btnExpandirDiferencias.Click, Sub() ExpandirRamasDiferencias()

        ' Orden de adición: con DockStyle.Right, el primero añadido va al borde derecho exterior
        pnlTreeToolbar.Controls.AddRange({_chkSoloDiferencias, _btnColapsarIdenticas, _btnExpandirDiferencias})

        _treeView = New TreeView With {
            .Dock = DockStyle.Fill,
            .BorderStyle = BorderStyle.FixedSingle,
            .Font = New Font("Segoe UI", 9.0F),
            .ShowLines = True,
            .ShowPlusMinus = True,
            .ShowRootLines = True,
            .FullRowSelect = False,
            .HideSelection = False,
            .ItemHeight = 22
        }
        AddHandler _treeView.AfterSelect, AddressOf OnTreeAfterSelect

        pnlTreeHost.Controls.Add(_treeView)
        pnlTreeHost.Controls.Add(pnlTreeToolbar)

        ' Divisor vertical estándar redimensionable con el ratón
        Dim splitterDivider As New Splitter With {
            .Dock = DockStyle.Left,
            .Width = 6,
            .BackColor = Color.FromArgb(226, 232, 240),
            .Cursor = Cursors.VSplit,
            .MinSize = 430,
            .MinExtra = 300
        }

        ' ==============================================================
        ' PANEL DERECHO: Metadata y Detalle de Inspección Interna
        ' ==============================================================
        Dim pnlDetailsHost As New Panel With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(10, 10, 20, 5)
        }

        Dim pnlMetaHeader As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 160,
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle,
            .Padding = New Padding(12, 10, 12, 10)
        }

        _lblDetalleTitulo = New Label With {
            .Text = "Seleccione una rama o archivo del árbol para ver su información",
            .Dock = DockStyle.Top,
            .Height = 22,
            .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(15, 23, 42)
        }

        _lblDetalleEstado = New Label With {
            .Text = "Estado: -",
            .Dock = DockStyle.Top,
            .Height = 22,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(71, 85, 105)
        }

        _lblDetalleMetadatos = New Label With {
            .Text = "Fechas y Tamaños: -",
            .Dock = DockStyle.Top,
            .Height = 44,
            .Font = New Font("Segoe UI", 8.5F),
            .ForeColor = Color.FromArgb(51, 65, 85)
        }

        _lblDetalleRutas = New Label With {
            .Text = "Rutas: -",
            .Dock = DockStyle.Fill,
            .Font = New Font("Consolas", 8.0F),
            .ForeColor = Color.FromArgb(100, 116, 139)
        }

        pnlMetaHeader.Controls.AddRange({_lblDetalleRutas, _lblDetalleMetadatos, _lblDetalleEstado, _lblDetalleTitulo})

        Dim lblDetallesTitle As New Label With {
            .Text = "Inspección Interna de Cambios (Materiales .adsklib / .asdklib, Código iLogic, XML):",
            .Dock = DockStyle.Top,
            .Height = 26,
            .Padding = New Padding(0, 6, 0, 0),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(51, 65, 85)
        }

        _txtDetallesInternos = New TextBox With {
            .Dock = DockStyle.Fill,
            .Multiline = True,
            .ReadOnly = True,
            .ScrollBars = ScrollBars.Both,
            .BackColor = Color.FromArgb(248, 250, 252),
            .Font = New Font("Consolas", 9.0F),
            .ForeColor = Color.FromArgb(30, 41, 59)
        }

        pnlDetailsHost.Controls.Add(_txtDetallesInternos)
        pnlDetailsHost.Controls.Add(lblDetallesTitle)
        pnlDetailsHost.Controls.Add(pnlMetaHeader)
        pnlCentral.Controls.AddRange({pnlDetailsHost, splitterDivider, pnlTreeHost})

        ' Bottom Footer Panel
        Dim pnlBottom As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 56,
            .BackColor = Color.FromArgb(241, 245, 249),
            .Padding = New Padding(20, 8, 20, 8)
        }

        _lblResumen = New Label With {
            .Text = "Haga clic en 'Comprobar Diferencias' para escanear en paralelo a máxima velocidad.",
            .Dock = DockStyle.Left,
            .AutoSize = True,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Italic),
            .ForeColor = Color.FromArgb(71, 85, 105),
            .Location = New Point(20, 18)
        }

        _progressBar = New ProgressBar With {
            .Dock = DockStyle.None,
            .Width = 200,
            .Height = 16,
            .Location = New Point(680, 20),
            .Style = ProgressBarStyle.Marquee,
            .Visible = False
        }

        _btnCerrar = New Button With {
            .Text = "Cerrar",
            .Dock = DockStyle.Right,
            .Width = 100,
            .BackColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        AddHandler _btnCerrar.Click, Sub() Me.Close()

        pnlBottom.Controls.AddRange({_lblResumen, _progressBar, _btnCerrar})

        Me.Controls.AddRange({pnlCentral, pnlConfig, pnlBottom, pnlHeader})
    End Sub

    Private Sub CargarConfiguracion()
        Dim cfg = AddinSettings.Current
        _txtRutaLocal.Text = If(Not String.IsNullOrEmpty(cfg.RutaLocalBibliotecas), cfg.RutaLocalBibliotecas, LibraryBackupSyncManager.DEFAULT_LOCAL_PATH)
        _txtRutaServidor.Text = If(Not String.IsNullOrEmpty(cfg.RutaServidorBibliotecas), cfg.RutaServidorBibliotecas, LibraryBackupSyncManager.DEFAULT_SERVER_PATH)

        Dim idx = _cmbIntervalo.Items.IndexOf(cfg.IntervaloBackupBibliotecas)
        _cmbIntervalo.SelectedIndex = If(idx >= 0, idx, 0)
        _chkAutoSync.Checked = cfg.AutoSyncBibliotecasAlIniciar

        VerificarEstadoRed()
    End Sub

    Private Sub GuardarConfiguracion()
        Dim cfg = AddinSettings.Current
        cfg.RutaLocalBibliotecas = _txtRutaLocal.Text.Trim()
        cfg.RutaServidorBibliotecas = _txtRutaServidor.Text.Trim()
        cfg.IntervaloBackupBibliotecas = _cmbIntervalo.SelectedItem.ToString()
        cfg.AutoSyncBibliotecasAlIniciar = _chkAutoSync.Checked
        AddinSettings.Save()
    End Sub

    Private Sub VerificarEstadoRed()
        Dim serverPath = _txtRutaServidor.Text.Trim()
        If Directory.Exists(serverPath) Then
            _lblEstadoRed.Text = "● Servidor de Red Conectado (5GbE)"
            _lblEstadoRed.ForeColor = Color.LimeGreen
        Else
            _lblEstadoRed.Text = "● Servidor No Accesible (Ruta no encontrada)"
            _lblEstadoRed.ForeColor = Color.Crimson
        End If
    End Sub

    Private Sub OnExaminarLocalClicked(sender As Object, e As EventArgs)
        Using fbd As New FolderBrowserDialog With {.SelectedPath = _txtRutaLocal.Text}
            If fbd.ShowDialog() = DialogResult.OK Then
                _txtRutaLocal.Text = fbd.SelectedPath
                GuardarConfiguracion()
            End If
        End Using
    End Sub

    Private Sub OnExaminarServidorClicked(sender As Object, e As EventArgs)
        Using fbd As New FolderBrowserDialog With {.SelectedPath = _txtRutaServidor.Text}
            If fbd.ShowDialog() = DialogResult.OK Then
                _txtRutaServidor.Text = fbd.SelectedPath
                GuardarConfiguracion()
                VerificarEstadoRed()
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Ejecución Asíncrona Ultra Rápida con Árbol Inteligente:
    ''' - Agrupa los archivos por carpetas y ramas.
    ''' - Si una rama es 100% idéntica, la mantiene colapsada (cerrada).
    ''' - Si una rama contiene diferencias, se despliega automáticamente y salta al archivo con cambios.
    ''' </summary>
    Private Async Sub OnCompararClicked(sender As Object, e As EventArgs)
        If _estaOcupado Then Exit Sub
        _estaOcupado = True

        Cursor = Cursors.WaitCursor
        _btnComparar.Enabled = False
        _btnPull.Enabled = False
        _btnPush.Enabled = False
        _progressBar.Visible = True

        GuardarConfiguracion()
        VerificarEstadoRed()

        Dim sw = Stopwatch.StartNew()

        Dim progreso As New Progress(Of String)(
            Sub(msg)
                _lblResumen.Text = msg
            End Sub)

        Try
            Dim localRoot = _txtRutaLocal.Text.Trim()
            Dim serverRoot = _txtRutaServidor.Text.Trim()

            ' Llamada explícita tipada a Task.Run(Of List(Of FileComparisonItem)) para evitar sobrecargas ambiguas en VB.NET
            _itemsComparados = Await Task.Run(Of List(Of FileComparisonItem))(
                Function() As List(Of FileComparisonItem)
                    Return _syncManager.CompararArchivos(localRoot, serverRoot, progreso)
                End Function)

            sw.Stop()
            Dim segs = sw.Elapsed.TotalSeconds

            ' Refinar en el hilo UI cualquier biblioteca .adsklib usando la API nativa de Inventor (STA)
            For Each item In _itemsComparados
                If item.Status <> FileSyncStatusEnum.Identical Then
                    Dim ext = Path.GetExtension(item.RelativePath).ToLowerInvariant()
                    If ext = ".adsklib" OrElse ext = ".asdklib" Then
                        _syncManager.RefinarInspeccionConInventor(item)
                    End If
                End If
            Next

            ' Construir el árbol jerárquico inteligente
            PoblarArbolJerarquico()

            Dim countPull = _itemsComparados.FindAll(Function(x) x.Status = FileSyncStatusEnum.ServerNewer OrElse x.Status = FileSyncStatusEnum.OnlyOnServer).Count
            Dim countPush = _itemsComparados.FindAll(Function(x) x.Status = FileSyncStatusEnum.LocalNewer OrElse x.Status = FileSyncStatusEnum.OnlyOnLocal).Count
            Dim countIdentical = _itemsComparados.FindAll(Function(x) x.Status = FileSyncStatusEnum.Identical).Count

            _btnPull.Enabled = (countPull > 0)
            _btnPush.Enabled = (countPush > 0)

            _lblResumen.Text = "⚡ " & _itemsComparados.Count.ToString() & " archivos analizados en " & segs.ToString("F2") & " seg | " & countPull.ToString() & " requieren PULL | " & countPush.ToString() & " requieren PUSH | " & countIdentical.ToString() & " idénticos"
        Catch ex As Exception
            MessageBox.Show("Error al comparar archivos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            _lblResumen.Text = "Error durante el escaneo."
        Finally
            _progressBar.Visible = False
            _btnComparar.Enabled = True
            _estaOcupado = False
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' Estructura auxiliar para resumir el estado de una rama/carpeta del árbol.
    ''' </summary>
    Private Class FolderNodeSummary
        Public Property FolderName As String
        Public Property RelativeFolderPath As String
        Public Property TotalFiles As Integer = 0
        Public Property IdenticalFiles As Integer = 0
        Public Property PullFiles As Integer = 0
        Public Property PushFiles As Integer = 0
        Public Property ConflictFiles As Integer = 0

        Public ReadOnly Property IsFullySynchronized As Boolean
            Get
                Return TotalFiles > 0 AndAlso IdenticalFiles = TotalFiles
            End Get
        End Property

        Public ReadOnly Property HasDifferences As Boolean
            Get
                Return (PullFiles > 0 OrElse PushFiles > 0 OrElse ConflictFiles > 0)
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Construye el árbol jerárquico por carpetas y ramas:
    ''' 1. Colapsa automáticamente las ramas donde todos los archivos son idénticos.
    ''' 2. Auto-despliega las ramas con diferencias.
    ''' 3. Selecciona automáticamente el primer archivo con cambios para que el usuario vea de inmediato qué cambió.
    ''' </summary>
    Private Sub PoblarArbolJerarquico()
        _treeView.BeginUpdate()
        _treeView.Nodes.Clear()

        Dim soloDiferencias = _chkSoloDiferencias.Checked

        ' Diccionario para encontrar o crear nodos de carpetas rápidamente
        Dim folderNodes As New Dictionary(Of String, TreeNode)(StringComparer.OrdinalIgnoreCase)
        Dim primerNodoDiferente As TreeNode = Nothing

        For Each item In _itemsComparados
            ' Filtro de solo diferencias si está activo
            If soloDiferencias AndAlso item.Status = FileSyncStatusEnum.Identical Then
                Continue For
            End If

            ' Partir la ruta relativa en partes de directorios y archivo
            Dim parts = item.RelativePath.Split(New Char() {"\"c, "/"c}, StringSplitOptions.RemoveEmptyEntries)
            Dim parentNodeCollection As TreeNodeCollection = _treeView.Nodes
            Dim currentPathAccum As String = ""

            ' Crear o navegar ramas de carpetas
            For i As Integer = 0 To parts.Length - 2
                Dim folderName = parts(i)
                currentPathAccum = If(String.IsNullOrEmpty(currentPathAccum), folderName, currentPathAccum & "\" & folderName)

                Dim folderNode As TreeNode = Nothing
                If Not folderNodes.TryGetValue(currentPathAccum, folderNode) Then
                    folderNode = New TreeNode(folderName) With {
                        .Tag = New FolderNodeSummary With {
                            .FolderName = folderName,
                            .RelativeFolderPath = currentPathAccum
                        }
                    }
                    parentNodeCollection.Add(folderNode)
                    folderNodes(currentPathAccum) = folderNode
                End If

                parentNodeCollection = folderNode.Nodes
            Next

            ' Crear nodo del archivo
            Dim fileName = parts(parts.Length - 1)
            Dim fileNode As New TreeNode() With {
                .Tag = item
            }

            Select Case item.Status
                Case FileSyncStatusEnum.LocalNewer
                    fileNode.Text = "⬆ " & fileName & "  [PUSH - Local más nuevo]"
                    fileNode.ForeColor = Color.FromArgb(194, 65, 12) ' Naranja fuerte
                    fileNode.NodeFont = New Font(_treeView.Font, FontStyle.Bold)
                    If primerNodoDiferente Is Nothing Then primerNodoDiferente = fileNode

                Case FileSyncStatusEnum.OnlyOnLocal
                    fileNode.Text = "➕ " & fileName & "  [PUSH - Solo en Local]"
                    fileNode.ForeColor = Color.FromArgb(217, 119, 6)
                    fileNode.NodeFont = New Font(_treeView.Font, FontStyle.Bold)
                    If primerNodoDiferente Is Nothing Then primerNodoDiferente = fileNode

                Case FileSyncStatusEnum.ServerNewer
                    fileNode.Text = "⬇ " & fileName & "  [PULL - Servidor más nuevo]"
                    fileNode.ForeColor = Color.FromArgb(37, 99, 235) ' Azul fuerte
                    fileNode.NodeFont = New Font(_treeView.Font, FontStyle.Bold)
                    If primerNodoDiferente Is Nothing Then primerNodoDiferente = fileNode

                Case FileSyncStatusEnum.OnlyOnServer
                    fileNode.Text = "📥 " & fileName & "  [PULL - Solo en Servidor]"
                    fileNode.ForeColor = Color.FromArgb(79, 70, 229)
                    fileNode.NodeFont = New Font(_treeView.Font, FontStyle.Bold)
                    If primerNodoDiferente Is Nothing Then primerNodoDiferente = fileNode

                Case FileSyncStatusEnum.Conflict
                    fileNode.Text = "⚡ " & fileName & "  [CONFLICTO]"
                    fileNode.ForeColor = Color.Crimson
                    fileNode.NodeFont = New Font(_treeView.Font, FontStyle.Bold)
                    If primerNodoDiferente Is Nothing Then primerNodoDiferente = fileNode

                Case FileSyncStatusEnum.Identical
                    fileNode.Text = "✔ " & fileName & "  [Idéntico]"
                    fileNode.ForeColor = Color.FromArgb(71, 85, 105)
            End Select

            parentNodeCollection.Add(fileNode)
        Next

        ' Calcular estadísticas recursivas de cada nodo de carpeta y aplicar auto-despliegue / colapso
        For Each rootNode As TreeNode In _treeView.Nodes
            ActualizarEstadisticasRama(rootNode)
        Next

        _treeView.EndUpdate()

        ' REGLA DE ORO SOLICITADA POR EL USUARIO:
        ' Si hay archivos con diferencias, desplegar la rama y llevar el foco automáticamente al primer archivo diferente
        ' Si todo es idéntico, mantener colapsado y seleccionar la raíz sin estorbar la vista.
        If primerNodoDiferente IsNot Nothing Then
            _treeView.SelectedNode = primerNodoDiferente
            primerNodoDiferente.EnsureVisible()
            _treeView.Focus()
        ElseIf _treeView.Nodes.Count > 0 Then
            _treeView.SelectedNode = _treeView.Nodes(0)
            _treeView.Nodes(0).Collapse()
        End If
    End Sub

    ''' <summary>
    ''' Recorre recursivamente una rama de carpetas acumulando contadores de archivos y aplicando la lógica visual:
    ''' - Si todos los archivos de la rama son idénticos: Se colapsa y muestra marca verde ✔ (100% Sincronizado).
    ''' - Si contiene diferencias: Se resalta en negrita/naranja ⚠️ y se AUTO-EXPANDE.
    ''' </summary>
    Private Function ActualizarEstadisticasRama(node As TreeNode) As FolderNodeSummary
        If TypeOf node.Tag Is FileComparisonItem Then
            Dim fItem = DirectCast(node.Tag, FileComparisonItem)
            Dim s As New FolderNodeSummary()
            s.TotalFiles = 1
            Select Case fItem.Status
                Case FileSyncStatusEnum.Identical
                    s.IdenticalFiles = 1
                Case FileSyncStatusEnum.LocalNewer, FileSyncStatusEnum.OnlyOnLocal
                    s.PushFiles = 1
                Case FileSyncStatusEnum.ServerNewer, FileSyncStatusEnum.OnlyOnServer
                    s.PullFiles = 1
                Case FileSyncStatusEnum.Conflict
                    s.ConflictFiles = 1
            End Select
            Return s
        End If

        Dim folderSummary = TryCast(node.Tag, FolderNodeSummary)
        If folderSummary Is Nothing Then
            folderSummary = New FolderNodeSummary With {.FolderName = node.Text}
            node.Tag = folderSummary
        End If

        ' Reiniciar contadores para recálculo
        folderSummary.TotalFiles = 0
        folderSummary.IdenticalFiles = 0
        folderSummary.PullFiles = 0
        folderSummary.PushFiles = 0
        folderSummary.ConflictFiles = 0

        For Each childNode As TreeNode In node.Nodes
            Dim childSum = ActualizarEstadisticasRama(childNode)
            folderSummary.TotalFiles += childSum.TotalFiles
            folderSummary.IdenticalFiles += childSum.IdenticalFiles
            folderSummary.PullFiles += childSum.PullFiles
            folderSummary.PushFiles += childSum.PushFiles
            folderSummary.ConflictFiles += childSum.ConflictFiles
        Next

        ' Formatear la etiqueta de la carpeta y decidir si se expande o colapsa
        If folderSummary.IsFullySynchronized Then
            node.Text = "📁 " & folderSummary.FolderName & "  (" & folderSummary.TotalFiles.ToString() & " arch.)  ✔ 100% Sincronizado"
            node.ForeColor = Color.FromArgb(22, 101, 52) ' Verde legible
            node.NodeFont = New Font(_treeView.Font, FontStyle.Regular)
            ' Permanecer colapsado si es idéntico
            node.Collapse()
        ElseIf folderSummary.HasDifferences Then
            Dim diffDesc As New List(Of String)
            If folderSummary.PullFiles > 0 Then diffDesc.Add(folderSummary.PullFiles.ToString() & " PULL")
            If folderSummary.PushFiles > 0 Then diffDesc.Add(folderSummary.PushFiles.ToString() & " PUSH")
            If folderSummary.ConflictFiles > 0 Then diffDesc.Add(folderSummary.ConflictFiles.ToString() & " Conflictos")

            node.Text = "📁 " & folderSummary.FolderName & "  (" & folderSummary.TotalFiles.ToString() & " arch. — ⚠️ " & String.Join(", ", diffDesc) & ")"
            node.ForeColor = Color.FromArgb(194, 65, 12) ' Naranja alerta
            node.NodeFont = New Font(_treeView.Font, FontStyle.Bold)
            ' AUTO-DESPLEGAR la rama porque contiene cambios importantes
            node.Expand()
        Else
            node.Text = "📁 " & folderSummary.FolderName
            node.ForeColor = Color.FromArgb(71, 85, 105)
        End If

        Return folderSummary
    End Function

    Private Sub AplicarFiltroArbol()
        PoblarArbolJerarquico()
    End Sub

    Private Sub ColapsarRamasIdenticas()
        _treeView.BeginUpdate()
        For Each rootNode As TreeNode In _treeView.Nodes
            ColapsarRamasIdenticasRecursivo(rootNode)
        Next
        _treeView.EndUpdate()
    End Sub

    Private Sub ColapsarRamasIdenticasRecursivo(node As TreeNode)
        Dim folderSum = TryCast(node.Tag, FolderNodeSummary)
        If folderSum IsNot Nothing Then
            If folderSum.IsFullySynchronized Then
                node.Collapse(True)
            Else
                For Each child As TreeNode In node.Nodes
                    ColapsarRamasIdenticasRecursivo(child)
                Next
            End If
        End If
    End Sub

    Private Sub ExpandirRamasDiferencias()
        _treeView.BeginUpdate()
        For Each rootNode As TreeNode In _treeView.Nodes
            ExpandirRamasDiferenciasRecursivo(rootNode)
        Next
        _treeView.EndUpdate()
    End Sub

    Private Sub ExpandirRamasDiferenciasRecursivo(node As TreeNode)
        Dim folderSum = TryCast(node.Tag, FolderNodeSummary)
        If folderSum IsNot Nothing AndAlso folderSum.HasDifferences Then
            node.Expand()
            For Each child As TreeNode In node.Nodes
                ExpandirRamasDiferenciasRecursivo(child)
            Next
        End If
    End Sub

    ''' <summary>
    ''' Al seleccionar cualquier nodo del árbol (archivo o carpeta), muestra inmediatamente la información correspondiente.
    ''' </summary>
    Private Sub OnTreeAfterSelect(sender As Object, e As TreeViewEventArgs)
        If e.Node Is Nothing Then Exit Sub

        If TypeOf e.Node.Tag Is FileComparisonItem Then
            MostrarDetallesArchivo(DirectCast(e.Node.Tag, FileComparisonItem))
        ElseIf TypeOf e.Node.Tag Is FolderNodeSummary Then
            MostrarDetallesCarpeta(DirectCast(e.Node.Tag, FolderNodeSummary))
        End If
    End Sub

    Private Sub MostrarDetallesArchivo(item As FileComparisonItem)
        ' Resolver propietarios de forma diferida si no se habían cargado
        If item.LocalModifiedBy = "-" AndAlso File.Exists(item.LocalFullPath) Then
            item.LocalModifiedBy = _syncManager.ObtenerPropietarioLazy(item.LocalFullPath)
        End If
        If item.ServerModifiedBy = "-" AndAlso File.Exists(item.ServerFullPath) Then
            item.ServerModifiedBy = _syncManager.ObtenerPropietarioLazy(item.ServerFullPath)
        End If

        ' Si es biblioteca .adsklib con diferencias, asegurar inspección nativa actualizada con Inventor en hilo UI
        Dim extLib = Path.GetExtension(item.RelativePath).ToLowerInvariant()
        If (extLib = ".adsklib" OrElse extLib = ".asdklib") AndAlso item.Status <> FileSyncStatusEnum.Identical Then
            If item.InternalDetails.Count = 0 OrElse Not item.InternalDetails(0).Contains("AssetLibraries API") Then
                _syncManager.RefinarInspeccionConInventor(item)
            End If
        End If

        _lblDetalleTitulo.Text = "📄 " & Path.GetFileName(item.RelativePath) & " (" & item.RelativePath & ")"
        _lblDetalleEstado.Text = "Estado: " & item.StatusDescription

        Select Case item.Status
            Case FileSyncStatusEnum.LocalNewer, FileSyncStatusEnum.OnlyOnLocal
                _lblDetalleEstado.ForeColor = Color.FromArgb(194, 65, 12)
            Case FileSyncStatusEnum.ServerNewer, FileSyncStatusEnum.OnlyOnServer
                _lblDetalleEstado.ForeColor = Color.FromArgb(37, 99, 235)
            Case FileSyncStatusEnum.Identical
                _lblDetalleEstado.ForeColor = Color.FromArgb(22, 101, 52)
            Case Else
                _lblDetalleEstado.ForeColor = Color.Crimson
        End Select

        Dim localDateStr = If(item.LocalModified > DateTime.MinValue, item.LocalModified.ToString("yyyy-MM-dd HH:mm:ss"), "No existe")
        Dim serverDateStr = If(item.ServerModified > DateTime.MinValue, item.ServerModified.ToString("yyyy-MM-dd HH:mm:ss"), "No existe")
        Dim localSizeStr = If(item.LocalSize > 0, LibraryBackupSyncManager.FormatBytes(item.LocalSize), "0 B")
        Dim serverSizeStr = If(item.ServerSize > 0, LibraryBackupSyncManager.FormatBytes(item.ServerSize), "0 B")

        _lblDetalleMetadatos.Text = "• Local: " & localDateStr & " | Peso: " & localSizeStr & " | Usuario: " & item.LocalModifiedBy & vbCrLf &
                                   "• Servidor: " & serverDateStr & " | Peso: " & serverSizeStr & " | Usuario: " & item.ServerModifiedBy

        _lblDetalleRutas.Text = "Local:    " & item.LocalFullPath & vbCrLf &
                                "Servidor: " & item.ServerFullPath

        ' Mostrar Inspección de Contenido Interno
        Dim sb As New StringBuilder()
        sb.AppendLine("ARCHIVO: " & item.RelativePath)
        sb.AppendLine("Estado: " & item.StatusDescription)
        sb.AppendLine("Ruta Local:    " & item.LocalFullPath)
        sb.AppendLine("Ruta Servidor: " & item.ServerFullPath)
        sb.AppendLine("Modificado Local:    " & localDateStr & " (Usuario: " & item.LocalModifiedBy & ")")
        sb.AppendLine("Modificado Servidor: " & serverDateStr & " (Usuario: " & item.ServerModifiedBy & ")")
        sb.AppendLine()
        sb.AppendLine("DETALLE INTERNO DE CONTENIDO:")
        sb.AppendLine("--------------------------------------------------------------------------------")

        If item.InternalDetails.Count > 0 Then
            For Each d In item.InternalDetails
                sb.AppendLine(d)
            Next
        ElseIf item.Status = FileSyncStatusEnum.Identical Then
            sb.AppendLine("El archivo es 100% idéntico en tamaño y fecha entre local y servidor.")
        Else
            sb.AppendLine("Sin detalles internos adicionales.")
        End If

        _txtDetallesInternos.Text = sb.ToString()
    End Sub

    Private Sub MostrarDetallesCarpeta(sum As FolderNodeSummary)
        _lblDetalleTitulo.Text = "📁 Rama / Carpeta: " & sum.FolderName
        _lblDetalleRutas.Text = "Subruta: " & sum.RelativeFolderPath

        If sum.IsFullySynchronized Then
            _lblDetalleEstado.Text = "Estado Rama: ✔ 100% Sincronizada (" & sum.TotalFiles.ToString() & " archivos idénticos)"
            _lblDetalleEstado.ForeColor = Color.FromArgb(22, 101, 52)
        Else
            _lblDetalleEstado.Text = "Estado Rama: ⚠️ Con diferencias pendientes (" & sum.PullFiles.ToString() & " PULL, " & sum.PushFiles.ToString() & " PUSH)"
            _lblDetalleEstado.ForeColor = Color.FromArgb(194, 65, 12)
        End If

        _lblDetalleMetadatos.Text = "Total de archivos en esta rama: " & sum.TotalFiles.ToString() & vbCrLf &
                                   "• Idénticos: " & sum.IdenticalFiles.ToString() & " | PULL requeridos: " & sum.PullFiles.ToString() & " | PUSH pendientes: " & sum.PushFiles.ToString()

        Dim sb As New StringBuilder()
        sb.AppendLine("RESUMEN DE CARPETA: " & sum.RelativeFolderPath)
        sb.AppendLine("--------------------------------------------------------------------------------")
        sb.AppendLine("• Total de archivos analizados: " & sum.TotalFiles.ToString())
        sb.AppendLine("• Archivos sincronizados:       " & sum.IdenticalFiles.ToString())
        sb.AppendLine("• Archivos más nuevos en serv.: " & sum.PullFiles.ToString() & " (requieren PULL)")
        sb.AppendLine("• Archivos más nuevos en local: " & sum.PushFiles.ToString() & " (requieren PUSH)")
        sb.AppendLine("• Conflictos detectados:        " & sum.ConflictFiles.ToString())
        sb.AppendLine()
        If sum.IsFullySynchronized Then
            sb.AppendLine("Esta rama está completamente al día entre el equipo local y el servidor de red.")
            sb.AppendLine("Se mantiene colapsada en la vista para no generar ruido visual.")
        Else
            sb.AppendLine("La rama contiene archivos que no coinciden con el servidor.")
            sb.AppendLine("El árbol la ha desplegado automáticamente. Seleccione cada archivo individual para ver las diferencias internas.")
        End If

        _txtDetallesInternos.Text = sb.ToString()
    End Sub

    Private Async Sub OnPullClicked(sender As Object, e As EventArgs)
        Dim countPull = _itemsComparados.FindAll(Function(x) x.Status = FileSyncStatusEnum.ServerNewer OrElse x.Status = FileSyncStatusEnum.OnlyOnServer).Count
        If countPull = 0 Then Exit Sub

        Dim resp = MessageBox.Show(
            "¿Desea descargar y actualizar " & countPull.ToString() & " archivo(s) del servidor a su carpeta local (" & _txtRutaLocal.Text & ")?" & vbCrLf & vbCrLf &
            "• Transferencia multihilo optimizada para red 5GbE." & vbCrLf &
            "• Los archivos locales más antiguos serán reemplazados por la versión del servidor." & vbCrLf &
            "• Se registrará un log detallado en el servidor.",
            "Confirmar PULL (Servidor -> Local)",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question)

        If resp <> DialogResult.Yes Then Exit Sub

        Cursor = Cursors.WaitCursor
        _btnPull.Enabled = False
        _btnPush.Enabled = False
        _progressBar.Visible = True

        Dim progreso As New Progress(Of String)(Sub(m) _lblResumen.Text = m)

        Try
            Dim localRoot = _txtRutaLocal.Text.Trim()
            Dim serverRoot = _txtRutaServidor.Text.Trim()

            ' Llamada explícita tipada a Task.Run(Of Integer)
            Dim n As Integer = Await Task.Run(Of Integer)(
                Function() As Integer
                    Return _syncManager.EjecutarPull(_itemsComparados, localRoot, serverRoot, progreso)
                End Function)

            AddinSettings.Current.UltimaSincronizacionBibliotecas = DateTime.Now
            AddinSettings.Save()

            MessageBox.Show("PULL Completado: Se actualizaron " & n.ToString() & " archivo(s) desde el servidor a alta velocidad.", "PULL Completado", MessageBoxButtons.OK, MessageBoxIcon.Information)
            OnCompararClicked(Nothing, Nothing)
        Catch ex As Exception
            MessageBox.Show("Error durante PULL: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _progressBar.Visible = False
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Async Sub OnPushClicked(sender As Object, e As EventArgs)
        Dim countPush = _itemsComparados.FindAll(Function(x) x.Status = FileSyncStatusEnum.LocalNewer OrElse x.Status = FileSyncStatusEnum.OnlyOnLocal).Count
        If countPush = 0 Then Exit Sub

        Dim resp = MessageBox.Show(
            "¿Desea subir " & countPush.ToString() & " archivo(s) modificados en su equipo local al servidor (" & _txtRutaServidor.Text & ")?" & vbCrLf & vbCrLf &
            "• Transferencia multihilo con búfer ampliado para saturar red 5GbE." & vbCrLf &
            "• Se generará un reporte de cambios en el servidor para que los demás usuarios sepan qué novedades hay (ej. materiales agregados/eliminados)." & vbCrLf &
            "• Se actualizará el CHANGELOG_GENERAL.txt en el servidor.",
            "Confirmar PUSH (Local -> Servidor)",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question)

        If resp <> DialogResult.Yes Then Exit Sub

        Cursor = Cursors.WaitCursor
        _btnPull.Enabled = False
        _btnPush.Enabled = False
        _progressBar.Visible = True

        Dim progreso As New Progress(Of String)(Sub(m) _lblResumen.Text = m)

        Try
            Dim localRoot = _txtRutaLocal.Text.Trim()
            Dim serverRoot = _txtRutaServidor.Text.Trim()

            ' Llamada explícita tipada a Task.Run(Of Integer)
            Dim n As Integer = Await Task.Run(Of Integer)(
                Function() As Integer
                    Return _syncManager.EjecutarPush(_itemsComparados, localRoot, serverRoot, progreso)
                End Function)

            AddinSettings.Current.UltimaSincronizacionBibliotecas = DateTime.Now
            AddinSettings.Save()

            MessageBox.Show("PUSH Completado: Se subieron " & n.ToString() & " archivo(s) al servidor con registro de cambios.", "PUSH Completado", MessageBoxButtons.OK, MessageBoxIcon.Information)
            OnCompararClicked(Nothing, Nothing)
        Catch ex As Exception
            MessageBox.Show("Error durante PUSH: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _progressBar.Visible = False
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub OnAbrirLogsClicked(sender As Object, e As EventArgs)
        Dim serverLogs = Path.Combine(_txtRutaServidor.Text.Trim(), "Sync_Logs")
        If Directory.Exists(serverLogs) Then
            Try
                Process.Start("explorer.exe", serverLogs)
            Catch
            End Try
        Else
            Dim localLogs = Path.Combine(_txtRutaLocal.Text.Trim(), "Sync_Logs")
            If Directory.Exists(localLogs) Then
                Try : Process.Start("explorer.exe", localLogs) : Catch : End Try
            Else
                MessageBox.Show("Aún no se han generado logs de sincronización.", "Logs", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub

End Class
