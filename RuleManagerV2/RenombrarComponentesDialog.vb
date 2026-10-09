Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Imports IOPath = System.IO.Path
Imports IOFile = System.IO.File
Imports Color = System.Drawing.Color
Imports Point = System.Drawing.Point

''' <summary>
''' Representa un elemento (pieza o ensamblaje) para la revisión y asignación de nombres por Part Number.
''' </summary>
Public Class ComponenteRenombrarItem
    Public Property Occurrence As ComponentOccurrence
    Public Property Document As Document
    Public Property DocumentPath As String
    Public Property FileName As String
    Public Property Extension As String
    Public Property IsAssembly As Boolean
    Public Property IsRootAssembly As Boolean
    Public Property StockNumber As String
    Public Property PartNumber As String
    Public Property SuggestedName As String
    Public Property OriginalSuggestedName As String
    Public Property Thumbnail As Image

    Public ReadOnly Property ColumnaPiezaStock As String
        Get
            Dim tipoStr As String = If(IsRootAssembly, "[Raíz]", If(IsAssembly, "[Subensamblaje]", "[Pieza]"))
            If Not String.IsNullOrEmpty(StockNumber) Then
                Return StockNumber & " " & tipoStr
            Else
                Return "(Sin Stock) " & tipoStr
            End If
        End Get
    End Property

    Public ReadOnly Property SuggestedFileName As String
        Get
            Dim sNom As String = If(SuggestedName, "").Trim()
            If sNom.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) OrElse sNom.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then
                sNom = IOPath.GetFileNameWithoutExtension(sNom)
            End If
            Return sNom & Extension
        End Get
    End Property

    Public ReadOnly Property TieneCambio As Boolean
        Get
            Dim sNom As String = If(SuggestedName, "").Trim()
            If sNom.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) OrElse sNom.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then
                sNom = IOPath.GetFileNameWithoutExtension(sNom)
            End If
            Dim actualSinExt As String = IOPath.GetFileNameWithoutExtension(FileName)
            Return Not String.Equals(sNom, actualSinExt, StringComparison.OrdinalIgnoreCase)
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return If(Not String.IsNullOrEmpty(PartNumber), PartNumber, FileName) & " (" & FileName & ")"
    End Function
End Class

''' <summary>
''' Ventana interactiva que muestra las piezas y ensamblajes en celdas,
''' permitiendo editar el nombre sugerido (Part Number), hacer zoom en el modelo 3D,
''' abrir el archivo en Inventor, cancelar o aplicar los cambios al renombrador.
''' </summary>
Public Class RenombrarComponentesDialog
    Inherits Form

    <DllImport("user32.dll")>
    Private Shared Function EnableWindow(hWnd As IntPtr, bEnable As Boolean) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function

    Private _invApp As Inventor.Application
    Private _items As List(Of ComponenteRenombrarItem)
    Private _prefix As String
    Private _highlightSet As HighlightSet = Nothing

    Private _grid As DataGridView
    Private _cmbPiezas As ComboBox
    Private _picPreview As PictureBox
    Private _lblTipo As Label
    Private _lblStockNumber As Label
    Private _lblPartNumber As Label
    Private _lblActual As Label
    Private _lblNuevo As Label
    Private _lblResumen As Label
    Private _btnZoom As Button
    Private _btnAbrir As Button
    Private _btnRestaurar As Button
    Private _btnAplicar As Button
    Private _btnCancelar As Button
    Private _btnAplicarLateral As Button
    Private _btnCancelarLateral As Button
    Private _contextMenu As ContextMenuStrip

    Public ReadOnly Property Items As List(Of ComponenteRenombrarItem)
        Get
            Return _items
        End Get
    End Property

    Public Sub New(invApp As Inventor.Application, items As List(Of ComponenteRenombrarItem), prefix As String)
        _invApp = invApp
        _items = items
        _prefix = prefix
        InitializeComponent()
        CargarDatos()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Renombrar Componentes - Asignación de Nombres por Part Number"
        Me.Size = New Size(1020, 640)
        Me.MinimumSize = New Size(850, 520)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Font = New Font("Segoe UI", 9.0F)

        ' =========================================================
        ' PANEL SUPERIOR: Título, instrucciones y selector rápido
        ' =========================================================
        Dim pnlTop As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 72,
            .Padding = New Padding(12, 8, 12, 8),
            .BackColor = Color.FromArgb(248, 250, 252)
        }

        Dim lblTitulo As New Label With {
            .Text = "Revisión y Asignación de Nombres a Componentes",
            .Font = New Font("Segoe UI", 10.0F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(15, 23, 42),
            .Location = New Point(12, 8),
            .AutoSize = True
        }

        Dim lblDesc As New Label With {
            .Text = "Verifique los nombres sugeridos basados en Part Number. Puede editar el campo 'Nombre Sugerido' directamente en la celda." & vbCrLf & "Haga doble clic en una fila o use los botones de la derecha para hacer zoom o abrir la pieza.",
            .ForeColor = Color.FromArgb(100, 116, 139),
            .Location = New Point(12, 28),
            .AutoSize = True
        }

        _btnRestaurar = New Button With {
            .Text = "🔄 Restaurar Sugeridos Originales",
            .Location = New Point(680, 10),
            .Size = New Size(220, 26),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .BackColor = Color.White,
            .Font = New Font("Segoe UI", 8.5F)
        }
        AddHandler _btnRestaurar.Click, AddressOf OnRestaurarSugeridos

        Dim lblCmb As New Label With {
            .Text = "Buscar / Seleccionar:",
            .Location = New Point(530, 43),
            .AutoSize = True,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        }

        _cmbPiezas = New ComboBox With {
            .Location = New Point(680, 40),
            .Size = New Size(310, 24),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        AddHandler _cmbPiezas.SelectedIndexChanged, AddressOf OnCmbSelectedIndexChanged

        pnlTop.Controls.Add(lblTitulo)
        pnlTop.Controls.Add(lblDesc)
        pnlTop.Controls.Add(_btnRestaurar)
        pnlTop.Controls.Add(lblCmb)
        pnlTop.Controls.Add(_cmbPiezas)

        ' =========================================================
        ' PANEL DERECHO: Vista Previa, Metadatos y Botones de Acción
        ' =========================================================
        Dim pnlRight As New Panel With {
            .Dock = DockStyle.Right,
            .Width = 260,
            .Padding = New Padding(10),
            .AutoScroll = True,
            .BackColor = Color.FromArgb(241, 245, 249)
        }

        Dim grpPreview As New GroupBox With {
            .Text = "Detalle del Componente",
            .Dock = DockStyle.Fill,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(30, 41, 59)
        }

        _picPreview = New PictureBox With {
            .Location = New Point(15, 22),
            .Size = New Size(210, 160),
            .SizeMode = PictureBoxSizeMode.Zoom,
            .BorderStyle = BorderStyle.FixedSingle,
            .BackColor = Color.White
        }

        _lblTipo = New Label With {
            .Location = New Point(15, 190),
            .Size = New Size(210, 18),
            .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(71, 85, 105),
            .Text = "Tipo: -"
        }

        _lblStockNumber = New Label With {
            .Location = New Point(15, 212),
            .Size = New Size(210, 32),
            .Font = New Font("Segoe UI", 8.0F),
            .Text = "Stock Number: -"
        }

        _lblPartNumber = New Label With {
            .Location = New Point(15, 246),
            .Size = New Size(210, 32),
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(2, 132, 199),
            .Text = "Part Number: -"
        }

        _lblActual = New Label With {
            .Location = New Point(15, 280),
            .Size = New Size(210, 32),
            .Font = New Font("Segoe UI", 8.0F),
            .ForeColor = Color.FromArgb(100, 116, 139),
            .Text = "Archivo actual: -"
        }

        _lblNuevo = New Label With {
            .Location = New Point(15, 314),
            .Size = New Size(210, 36),
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(16, 185, 129),
            .Text = "Nuevo archivo: -"
        }

        _btnZoom = New Button With {
            .Text = "🔍 Hacer Zoom",
            .Location = New Point(15, 358),
            .Size = New Size(210, 32),
            .BackColor = Color.White,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(30, 41, 59),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler _btnZoom.Click, AddressOf OnBtnZoomClick

        _btnAbrir = New Button With {
            .Text = "📂 Abrir Elemento",
            .Location = New Point(15, 396),
            .Size = New Size(210, 32),
            .BackColor = Color.White,
            .Font = New Font("Segoe UI", 9.0F),
            .ForeColor = Color.FromArgb(30, 41, 59),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler _btnAbrir.Click, AddressOf OnBtnAbrirClick

        Dim lblTip As New Label With {
            .Text = "💡 Tip: Doble clic en una fila para hacer zoom directamente.",
            .Location = New Point(15, 436),
            .Size = New Size(210, 36),
            .Font = New Font("Segoe UI", 7.5F, FontStyle.Italic),
            .ForeColor = Color.FromArgb(148, 163, 184)
        }

        ' Botones de acción directa en el panel lateral derecho
        _btnAplicarLateral = New Button With {
            .Text = "✔️ Aplicar y Renombrar",
            .Location = New Point(15, 478),
            .Size = New Size(210, 34),
            .BackColor = Color.FromArgb(2, 132, 199),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        _btnAplicarLateral.FlatAppearance.BorderSize = 0
        AddHandler _btnAplicarLateral.Click, AddressOf OnAplicar

        _btnCancelarLateral = New Button With {
            .Text = "✖️ Cancelar",
            .Location = New Point(15, 518),
            .Size = New Size(210, 30),
            .BackColor = Color.White,
            .ForeColor = Color.FromArgb(51, 65, 85),
            .Font = New Font("Segoe UI", 8.5F),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        _btnCancelarLateral.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        AddHandler _btnCancelarLateral.Click, AddressOf OnCancelar

        grpPreview.Controls.Add(_picPreview)
        grpPreview.Controls.Add(_lblTipo)
        grpPreview.Controls.Add(_lblStockNumber)
        grpPreview.Controls.Add(_lblPartNumber)
        grpPreview.Controls.Add(_lblActual)
        grpPreview.Controls.Add(_lblNuevo)
        grpPreview.Controls.Add(_btnZoom)
        grpPreview.Controls.Add(_btnAbrir)
        grpPreview.Controls.Add(lblTip)
        grpPreview.Controls.Add(_btnAplicarLateral)
        grpPreview.Controls.Add(_btnCancelarLateral)
        pnlRight.Controls.Add(grpPreview)

        ' =========================================================
        ' PANEL INFERIOR: Resumen y Botones Aplicar / Cancelar
        ' =========================================================
        Dim pnlBottom As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 56,
            .Padding = New Padding(16, 8, 16, 8),
            .BackColor = Color.FromArgb(241, 245, 249)
        }

        Dim sepBottom As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 1,
            .BackColor = Color.FromArgb(203, 213, 225)
        }
        pnlBottom.Controls.Add(sepBottom)

        _lblResumen = New Label With {
            .Location = New Point(16, 18),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(51, 65, 85),
            .Text = "Total componentes: 0"
        }

        ' Contenedor de botones alineado a la derecha de la barra inferior
        Dim pnlBotonesInferiores As New FlowLayoutPanel With {
            .Dock = DockStyle.Right,
            .FlowDirection = FlowDirection.RightToLeft,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .WrapContents = False,
            .Padding = New Padding(0, 10, 16, 8)
        }

        _btnCancelar = New Button With {
            .Text = "Cancelar",
            .Size = New Size(110, 34),
            .BackColor = Color.White,
            .ForeColor = Color.FromArgb(51, 65, 85),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand,
            .Margin = New Padding(10, 0, 0, 0)
        }
        _btnCancelar.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        AddHandler _btnCancelar.Click, AddressOf OnCancelar

        _btnAplicar = New Button With {
            .Text = "Aplicar y Renombrar",
            .Size = New Size(170, 34),
            .BackColor = Color.FromArgb(2, 132, 199),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand,
            .Margin = New Padding(0, 0, 0, 0)
        }
        _btnAplicar.FlatAppearance.BorderSize = 0
        AddHandler _btnAplicar.Click, AddressOf OnAplicar

        pnlBotonesInferiores.Controls.Add(_btnCancelar)
        pnlBotonesInferiores.Controls.Add(_btnAplicar)

        pnlBottom.Controls.Add(pnlBotonesInferiores)
        pnlBottom.Controls.Add(_lblResumen)

        AddHandler pnlBottom.Resize, Sub(s, e)
                                        _lblResumen.Top = (pnlBottom.ClientSize.Height - _lblResumen.Height) \ 2
                                    End Sub

        ' =========================================================
        ' CENTRO: DataGridView con las 4 columnas
        ' =========================================================
        _grid = New DataGridView With {
            .Dock = DockStyle.Fill,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .RowHeadersVisible = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .MultiSelect = False,
            .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.None,
            .AutoGenerateColumns = False,
            .EditMode = DataGridViewEditMode.EditOnEnter,
            .Font = New Font("Segoe UI", 8.5F)
        }

        ConfigurarColumnasGrid()
        ConfigurarMenuContextual()

        AddHandler _grid.CellValueChanged, AddressOf OnGridCellValueChanged
        AddHandler _grid.CurrentCellDirtyStateChanged, AddressOf OnGridCurrentCellDirtyStateChanged
        AddHandler _grid.SelectionChanged, AddressOf OnGridSelectionChanged
        AddHandler _grid.CellDoubleClick, AddressOf OnGridCellDoubleClick
        AddHandler _grid.CellMouseClick, AddressOf OnGridCellMouseClick

        ' Agregar paneles asegurando visibilidad y Z-Order correcto
        Me.Controls.Add(_grid)
        Me.Controls.Add(pnlRight)
        Me.Controls.Add(pnlTop)
        Me.Controls.Add(pnlBottom)

        pnlBottom.BringToFront()
        pnlTop.BringToFront()
        pnlRight.BringToFront()
        _grid.BringToFront()

        AddHandler Me.Shown, Sub(s, e) HabilitarVentanaInventor()
        AddHandler Me.Activated, Sub(s, e) HabilitarVentanaInventor()
        AddHandler Me.FormClosing, Sub(s, e)
                                       LimpiarResaltado()
                                       HabilitarVentanaInventor()
                                   End Sub

        Me.AcceptButton = _btnAplicar
        Me.CancelButton = _btnCancelar
    End Sub

    Private Sub ConfigurarColumnasGrid()
        ' Columna 1: Pieza o ensamblaje con su stock number
        Dim colStock As New DataGridViewTextBoxColumn With {
            .HeaderText = "Pieza / Ensamblaje (Stock Number)",
            .Name = "colStock",
            .ReadOnly = True,
            .Width = 240
        }

        ' Columna 2: Part Number
        Dim colPart As New DataGridViewTextBoxColumn With {
            .HeaderText = "Part Number",
            .Name = "colPart",
            .ReadOnly = True,
            .Width = 170
        }

        ' Columna 3: Nombre actual del archivo
        Dim colActual As New DataGridViewTextBoxColumn With {
            .HeaderText = "Nombre Actual",
            .Name = "colActual",
            .ReadOnly = True,
            .Width = 180
        }

        ' Columna 4: Nombre sugerido (editable)
        Dim colSugerido As New DataGridViewTextBoxColumn With {
            .HeaderText = "Nombre Sugerido (Editable)",
            .Name = "colSugerido",
            .ReadOnly = False,
            .Width = 220,
            .DefaultCellStyle = New DataGridViewCellStyle With {
                .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
                .BackColor = Color.FromArgb(254, 252, 232),
                .SelectionBackColor = Color.FromArgb(253, 224, 71),
                .SelectionForeColor = Color.Black
            }
        }

        _grid.Columns.AddRange(New DataGridViewColumn() {
            colStock, colPart, colActual, colSugerido
        })
    End Sub

    Private Sub ConfigurarMenuContextual()
        _contextMenu = New ContextMenuStrip()

        Dim mnuZoom As New ToolStripMenuItem("🔍 Hacer Zoom en el Modelo")
        AddHandler mnuZoom.Click, AddressOf OnBtnZoomClick
        _contextMenu.Items.Add(mnuZoom)

        Dim mnuAbrir As New ToolStripMenuItem("📂 Abrir Elemento en Inventor")
        AddHandler mnuAbrir.Click, AddressOf OnBtnAbrirClick
        _contextMenu.Items.Add(mnuAbrir)

        _contextMenu.Items.Add(New ToolStripSeparator())

        Dim mnuRestaurarFila As New ToolStripMenuItem("🔄 Restaurar a Part Number Original")
        AddHandler mnuRestaurarFila.Click, Sub(s, e)
                                               Dim item = GetSelectedItem()
                                               If item IsNot Nothing Then
                                                   item.SuggestedName = item.OriginalSuggestedName
                                                   If _grid.CurrentRow IsNot Nothing Then
                                                       _grid.CurrentRow.Cells("colSugerido").Value = item.SuggestedName
                                                   End If
                                                   ActualizarDetalleItem(item)
                                                   ActualizarResumen()
                                               End If
                                           End Sub
        _contextMenu.Items.Add(mnuRestaurarFila)

        _grid.ContextMenuStrip = _contextMenu
    End Sub

    Private Sub CargarDatos()
        _grid.Rows.Clear()
        _cmbPiezas.Items.Clear()

        For Each item In _items
            Dim rowIdx = _grid.Rows.Add(
                item.ColumnaPiezaStock,
                item.PartNumber,
                item.FileName,
                item.SuggestedName
            )
            _grid.Rows(rowIdx).Tag = item
            _cmbPiezas.Items.Add(item)
        Next

        If _cmbPiezas.Items.Count > 0 Then
            _cmbPiezas.SelectedIndex = 0
        End If

        ActualizarResumen()
    End Sub

    Private Function GetSelectedItem() As ComponenteRenombrarItem
        If _grid.CurrentRow IsNot Nothing AndAlso _grid.CurrentRow.Tag IsNot Nothing Then
            Return DirectCast(_grid.CurrentRow.Tag, ComponenteRenombrarItem)
        End If
        Return Nothing
    End Function

    Private Sub ActualizarDetalleItem(item As ComponenteRenombrarItem)
        If item Is Nothing Then
            _picPreview.Image = Nothing
            _lblTipo.Text = "Tipo: -"
            _lblStockNumber.Text = "Stock Number: -"
            _lblPartNumber.Text = "Part Number: -"
            _lblActual.Text = "Archivo actual: -"
            _lblNuevo.Text = "Nuevo archivo: -"
            Exit Sub
        End If

        _picPreview.Image = item.Thumbnail
        _lblTipo.Text = "Tipo: " & If(item.IsRootAssembly, "Ensamblaje Raíz", If(item.IsAssembly, "Subensamblaje (.iam)", "Pieza (.ipt)"))
        _lblStockNumber.Text = "Stock Number: " & If(String.IsNullOrEmpty(item.StockNumber), "(vacío)", item.StockNumber)
        _lblPartNumber.Text = "Part Number: " & If(String.IsNullOrEmpty(item.PartNumber), "(vacío)", item.PartNumber)
        _lblActual.Text = "Archivo actual: " & item.FileName
        _lblNuevo.Text = "Nuevo archivo: " & item.SuggestedFileName

        If item.TieneCambio Then
            _lblNuevo.ForeColor = Color.FromArgb(16, 185, 129)
        Else
            _lblNuevo.ForeColor = Color.FromArgb(100, 116, 139)
        End If
    End Sub

    Private Sub ActualizarResumen()
        Dim total As Integer = _items.Count
        Dim conCambio As Integer = 0
        Dim sinCambio As Integer = 0

        For Each it In _items
            If it.TieneCambio Then
                conCambio += 1
            Else
                sinCambio += 1
            End If
        Next

        _lblResumen.Text = "Total componentes: " & total & "  |  Con nuevo nombre: " & conCambio & "  |  Sin cambio: " & sinCambio
    End Sub

    ''' <summary>
    ''' Devuelve el mapeo de rutas originales a nombres finales configurados por el usuario.
    ''' </summary>
    Public Function ObtenerNombresPersonalizados() As Dictionary(Of String, String)
        Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        For Each it In _items
            If Not String.IsNullOrEmpty(it.DocumentPath) Then
                Dim baseNom As String = If(it.SuggestedName, "").Trim()
                If baseNom.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) OrElse baseNom.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then
                    baseNom = IOPath.GetFileNameWithoutExtension(baseNom)
                End If
                If Not dict.ContainsKey(it.DocumentPath) Then
                    dict.Add(it.DocumentPath, baseNom)
                End If
            End If
        Next
        Return dict
    End Function

    ' =========================================================
    ' INTERACCIÓN CON INVENTOR Y RESALTADO CELESTE
    ' =========================================================

    Public Sub HabilitarVentanaInventor()
        If _invApp IsNot Nothing Then
            Try
                Dim hwnd As New IntPtr(_invApp.MainFrameHWND)
                If hwnd <> IntPtr.Zero Then
                    EnableWindow(hwnd, True)
                End If
            Catch
            End Try
        End If
    End Sub

    Public Sub LimpiarResaltado()
        If _highlightSet IsNot Nothing Then
            Try
                _highlightSet.Clear()
            Catch
            End Try
            Try
                _highlightSet.Delete()
            Catch
            End Try
            _highlightSet = Nothing
            Try
                If _invApp IsNot Nothing AndAlso _invApp.ActiveView IsNot Nothing Then
                    _invApp.ActiveView.Update()
                End If
            Catch
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Agrega una ocurrencia, sus subocurrencias y todos sus cuerpos sólidos (SurfaceBodies)
    ''' al HighlightSet para asegurar que toda la geometría (de piezas o subensamblajes) se coloree en el visor 3D.
    ''' </summary>
    Private Sub AgregarOcurrenciaAHighlight(occ As ComponentOccurrence, hs As HighlightSet)
        If occ Is Nothing OrElse hs Is Nothing Then Return
        Try
            hs.AddItem(occ)
        Catch
        End Try

        Try
            If occ.SurfaceBodies IsNot Nothing Then
                For Each sb As SurfaceBody In occ.SurfaceBodies
                    Try
                        hs.AddItem(sb)
                    Catch
                    End Try
                Next
            End If
        Catch
        End Try

        Try
            Dim subOccs As Object = Nothing
            Try : subOccs = occ.SubOccurrences : Catch : End Try
            If subOccs IsNot Nothing Then
                Dim enumerable As System.Collections.IEnumerable = TryCast(subOccs, System.Collections.IEnumerable)
                If enumerable IsNot Nothing Then
                    For Each rawSub In enumerable
                        Dim subOcc As ComponentOccurrence = TryCast(rawSub, ComponentOccurrence)
                        If subOcc IsNot Nothing Then
                            Dim isSuppressed As Boolean = False
                            Try : isSuppressed = subOcc.Suppressed : Catch : End Try
                            If Not isSuppressed Then
                                AgregarOcurrenciaAHighlight(subOcc, hs)
                            End If
                        End If
                    Next
                End If
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Aplica una capa interactiva de color celeste vibrante sobre el componente en el visor 3D.
    ''' </summary>
    Public Sub ResaltarElemento(occ As ComponentOccurrence)
        If occ Is Nothing OrElse _invApp Is Nothing Then
            LimpiarResaltado()
            Return
        End If

        Try
            Dim doc As Document = _invApp.ActiveDocument
            If doc Is Nothing Then Return

            LimpiarResaltado()

            ' 1. Obtener o crear HighlightSet en el documento (ensamblaje)
            Dim asmDoc As AssemblyDocument = TryCast(doc, AssemblyDocument)
            If asmDoc IsNot Nothing Then
                Try
                    _highlightSet = asmDoc.HighlightSets.Add()
                Catch
                    Try
                        _highlightSet = asmDoc.HighlightSets.Item(1)
                    Catch
                    End Try
                End Try
            End If

            If _highlightSet Is Nothing Then
                Try
                    _highlightSet = doc.CreateHighlightSet()
                Catch
                End Try
            End If

            If _highlightSet IsNot Nothing Then
                ' Color celeste vibrante e interactivo (R: 0, G: 195, B: 255)
                Dim colorCeleste As Inventor.Color = _invApp.TransientObjects.CreateColor(0, 195, 255)
                Try
                    colorCeleste.Opacity = 0.85
                Catch
                End Try
                _highlightSet.Color = colorCeleste

                ' Resaltar la ocurrencia, sus subocurrencias y cuerpos sólidos 3D
                AgregarOcurrenciaAHighlight(occ, _highlightSet)
            End If

            ' Limpiar la selección de SelectSet para que no enmascare ni oculte el color celeste
            Try
                doc.SelectSet.Clear()
            Catch
            End Try

            If _invApp.ActiveView IsNot Nothing Then
                _invApp.ActiveView.Update()
            End If
        Catch
        End Try
    End Sub

    ' =========================================================
    ' ACCIONES: ZOOM Y ABRIR ELEMENTO
    ' =========================================================

    Public Sub HacerZoom(item As ComponenteRenombrarItem)
        If item Is Nothing OrElse _invApp Is Nothing Then Return
        Try
            If item.IsRootAssembly OrElse item.Occurrence Is Nothing Then
                LimpiarResaltado()
                If _invApp.ActiveView IsNot Nothing Then
                    _invApp.ActiveView.Fit()
                    _invApp.ActiveView.Update()
                End If
                Return
            End If

            Dim occ As ComponentOccurrence = item.Occurrence

            ' Si el documento activo actual no es el ensamblaje contenedor de la ocurrencia, activarlo
            Try
                Dim parentDoc As Document = Nothing
                If occ.Parent IsNot Nothing Then
                    Try : parentDoc = occ.Parent.Document : Catch : End Try
                End If
                If parentDoc IsNot Nothing AndAlso _invApp.ActiveDocument IsNot parentDoc Then
                    If parentDoc.Views IsNot Nothing AndAlso parentDoc.Views.Count > 0 Then
                        parentDoc.Views.Item(1).Activate()
                    Else
                        parentDoc.Activate()
                    End If
                End If
            Catch
            End Try

            Dim doc As Document = _invApp.ActiveDocument
            If doc Is Nothing Then Return

            ' Aplicar capa interactiva de color celeste (HighlightSet)
            ResaltarElemento(occ)

            Dim view As Inventor.View = _invApp.ActiveView
            If view IsNot Nothing Then
                Dim cam As Camera = view.Camera
                Dim box As Box = occ.RangeBox

                Dim centerX As Double = (box.MinPoint.X + box.MaxPoint.X) / 2.0
                Dim centerY As Double = (box.MinPoint.Y + box.MaxPoint.Y) / 2.0
                Dim centerZ As Double = (box.MinPoint.Z + box.MaxPoint.Z) / 2.0

                Dim centerPt As Inventor.Point = _invApp.TransientGeometry.CreatePoint(centerX, centerY, centerZ)

                Dim sizeX As Double = Math.Abs(box.MaxPoint.X - box.MinPoint.X)
                Dim sizeY As Double = Math.Abs(box.MaxPoint.Y - box.MinPoint.Y)
                Dim sizeZ As Double = Math.Abs(box.MaxPoint.Z - box.MinPoint.Z)
                Dim maxSize As Double = Math.Max(sizeX, Math.Max(sizeY, sizeZ))
                Dim distance As Double = Math.Max(maxSize * 2.2, 5.0)

                cam.Target = centerPt
                cam.Eye = _invApp.TransientGeometry.CreatePoint(
                    centerPt.X + distance,
                    centerPt.Y + distance,
                    centerPt.Z + distance)
                cam.UpVector = _invApp.TransientGeometry.CreateUnitVector(0, 0, 1)
                cam.Apply()
                view.Update()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub AbrirElemento(item As ComponenteRenombrarItem)
        If item Is Nothing OrElse _invApp Is Nothing Then Return

        Dim ruta As String = item.DocumentPath
        If String.IsNullOrEmpty(ruta) AndAlso item.Document IsNot Nothing Then
            Try
                ruta = item.Document.FullFileName
                If String.IsNullOrEmpty(ruta) Then ruta = item.Document.FullDocumentName
            Catch
            End Try
        End If

        If String.IsNullOrEmpty(ruta) AndAlso item.Occurrence IsNot Nothing AndAlso item.Occurrence.Definition IsNot Nothing Then
            Try
                Dim defDoc As Document = item.Occurrence.Definition.Document
                If defDoc IsNot Nothing Then
                    ruta = defDoc.FullFileName
                    If String.IsNullOrEmpty(ruta) Then ruta = defDoc.FullDocumentName
                End If
            Catch
            End Try
        End If

        If String.IsNullOrEmpty(ruta) Then
            MessageBox.Show("No se encontró la ruta del archivo del elemento seleccionado.", "Ruta no disponible", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If Not IOFile.Exists(ruta) Then
            MessageBox.Show("El archivo no existe en disco:" & vbCrLf & ruta, "Archivo no encontrado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Asegurar que Inventor quede completamente habilitado para interacción
        HabilitarVentanaInventor()

        Dim abiertoExitoso As Boolean = False

        ' ESTRATEGIA 1: Si el documento ya tiene una ventana visible abierta en Inventor, activarla
        Try
            For Each d As Document In _invApp.Documents
                If String.Equals(d.FullFileName, ruta, StringComparison.OrdinalIgnoreCase) Then
                    If d.Views IsNot Nothing AndAlso d.Views.Count > 0 Then
                        Try
                            d.Views.Item(1).Activate()
                            abiertoExitoso = True
                            Exit For
                        Catch
                            Try
                                d.Activate()
                                abiertoExitoso = True
                                Exit For
                            Catch
                            End Try
                        End Try
                    End If
                End If
            Next
        Catch
        End Try

        ' ESTRATEGIA 2: Abrir con la API de Inventor usando SilentOperation para evitar bloqueos de diálogos internos
        If Not abiertoExitoso Then
            Dim bSilentOriginal As Boolean = False
            Dim cambioSilent As Boolean = False
            Try
                Try
                    bSilentOriginal = _invApp.SilentOperation
                    _invApp.SilentOperation = True
                    cambioSilent = True
                Catch
                End Try

                Dim docAbierto As Document = _invApp.Documents.Open(ruta, True)
                If docAbierto IsNot Nothing Then
                    If docAbierto.Views IsNot Nothing AndAlso docAbierto.Views.Count > 0 Then
                        Try
                            docAbierto.Views.Item(1).Activate()
                        Catch
                        End Try
                    End If
                    abiertoExitoso = True
                End If
            Catch exCom As Exception
                ' Si falla la llamada directa de la API COM, pasamos a la Estrategia 3 como fallback
            Finally
                If cambioSilent Then
                    Try
                        _invApp.SilentOperation = bSilentOriginal
                    Catch
                    End Try
                End If
            End Try
        End If

        ' ESTRATEGIA 3: Fallback mediante Shell de Windows (Process.Start)
        If Not abiertoExitoso Then
            Try
                Dim psi As New System.Diagnostics.ProcessStartInfo() With {
                    .FileName = ruta,
                    .UseShellExecute = True
                }
                System.Diagnostics.Process.Start(psi)
                abiertoExitoso = True
            Catch exShell As Exception
                MessageBox.Show("No se pudo abrir el elemento: " & exShell.Message, "Error al abrir", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If

        ' Si se abrió con éxito, dar foco a la ventana principal de Inventor
        ' para que el usuario pueda rotar, inspeccionar o editar el elemento directamente
        ' sin necesidad de cerrar este diálogo.
        If abiertoExitoso Then
            HabilitarVentanaInventor()
            Try
                Dim hwnd As New IntPtr(_invApp.MainFrameHWND)
                If hwnd <> IntPtr.Zero Then
                    SetForegroundWindow(hwnd)
                End If
            Catch
            End Try
        End If
    End Sub

    ' =========================================================
    ' EVENTOS DE INTERFAZ
    ' =========================================================

    Private Sub OnBtnZoomClick(sender As Object, e As EventArgs)
        Dim item = GetSelectedItem()
        If item IsNot Nothing Then
            HacerZoom(item)
        End If
    End Sub

    Private Sub OnBtnAbrirClick(sender As Object, e As EventArgs)
        Dim item = GetSelectedItem()
        If item IsNot Nothing Then
            AbrirElemento(item)
        End If
    End Sub

    Private Sub OnGridCellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Exit Sub
        ' Si se hace doble clic sobre cualquier columna que no sea la editable, hacer zoom
        If e.ColumnIndex <> _grid.Columns("colSugerido").Index Then
            Dim item = GetSelectedItem()
            If item IsNot Nothing Then
                HacerZoom(item)
            End If
        End If
    End Sub

    Private Sub OnGridCellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs)
        If e.Button = MouseButtons.Right AndAlso e.RowIndex >= 0 Then
            _grid.ClearSelection()
            _grid.Rows(e.RowIndex).Selected = True
            _grid.CurrentCell = _grid.Rows(e.RowIndex).Cells(If(e.ColumnIndex >= 0, e.ColumnIndex, 0))
        End If
    End Sub

    Private Sub OnGridCurrentCellDirtyStateChanged(sender As Object, e As EventArgs)
        If _grid.IsCurrentCellDirty Then
            _grid.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub OnGridCellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex = _grid.Columns("colSugerido").Index Then
            Dim item = DirectCast(_grid.Rows(e.RowIndex).Tag, ComponenteRenombrarItem)
            If item IsNot Nothing Then
                Dim val = _grid.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                Dim nuevoTexto As String = If(val IsNot Nothing, val.ToString().Trim(), "")
                item.SuggestedName = nuevoTexto
                ActualizarDetalleItem(item)
                ActualizarResumen()
            End If
        End If
    End Sub

    Private Sub OnGridSelectionChanged(sender As Object, e As EventArgs)
        Dim item = GetSelectedItem()
        ActualizarDetalleItem(item)

        If item IsNot Nothing AndAlso _cmbPiezas.SelectedItem IsNot item Then
            _cmbPiezas.SelectedItem = item
        End If

        If item IsNot Nothing AndAlso item.Occurrence IsNot Nothing Then
            ResaltarElemento(item.Occurrence)
        Else
            LimpiarResaltado()
        End If
    End Sub

    Private Sub OnCmbSelectedIndexChanged(sender As Object, e As EventArgs)
        Dim item = DirectCast(_cmbPiezas.SelectedItem, ComponenteRenombrarItem)
        If item IsNot Nothing Then
            ActualizarDetalleItem(item)

            For Each row As DataGridViewRow In _grid.Rows
                If row.Tag Is item Then
                    If _grid.CurrentRow IsNot row Then
                        _grid.CurrentCell = row.Cells(0)
                    End If
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub OnRestaurarSugeridos(sender As Object, e As EventArgs)
        Dim res = MessageBox.Show("¿Desea restaurar todos los nombres sugeridos a su Part Number original?", "Confirmar restauración", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = DialogResult.Yes Then
            For Each row As DataGridViewRow In _grid.Rows
                Dim item = DirectCast(row.Tag, ComponenteRenombrarItem)
                If item IsNot Nothing Then
                    item.SuggestedName = item.OriginalSuggestedName
                    row.Cells("colSugerido").Value = item.SuggestedName
                End If
            Next
            ActualizarDetalleItem(GetSelectedItem())
            ActualizarResumen()
        End If
    End Sub

    Private Sub OnAplicar(sender As Object, e As EventArgs)
        ' Validar nombres sugeridos
        Dim invalidos As New List(Of String)()
        Dim vacios As New List(Of String)()
        Dim sInvalidChars As String = "\/:*?""<>|"

        For Each item In _items
            Dim nom = If(item.SuggestedName, "").Trim()
            If String.IsNullOrEmpty(nom) Then
                vacios.Add(item.FileName)
            Else
                For Each c In sInvalidChars
                    If nom.Contains(c) Then
                        invalidos.Add(item.FileName & " (contiene '" & c & "')")
                        Exit For
                    End If
                Next
            End If
        Next

        If vacios.Count > 0 Then
            Dim msg = "Los siguientes elementos tienen el nombre sugerido vacío y no podrán ser renombrados:" & vbCrLf & vbCrLf &
                      String.Join(vbCrLf, vacios.Take(10)) &
                      If(vacios.Count > 10, vbCrLf & "... y " & (vacios.Count - 10) & " más.", "") & vbCrLf & vbCrLf &
                      "¿Desea continuar de todas formas? (Los vacíos conservarán su nombre actual)."
            Dim dr = MessageBox.Show(msg, "Nombres vacíos", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If dr = DialogResult.No Then Return
        End If

        If invalidos.Count > 0 Then
            Dim msg = "Los siguientes elementos contienen caracteres no válidos para nombres de archivo (\ / : * ? "" < > |):" & vbCrLf & vbCrLf &
                      String.Join(vbCrLf, invalidos.Take(10)) &
                      If(invalidos.Count > 10, vbCrLf & "... y " & (invalidos.Count - 10) & " más.", "") & vbCrLf & vbCrLf &
                      "Por favor, corrija estos nombres antes de continuar."
            MessageBox.Show(msg, "Caracteres no permitidos", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        DialogResult = DialogResult.OK
        Close()
    End Sub

    Private Sub OnCancelar(sender As Object, e As EventArgs)
        DialogResult = DialogResult.Cancel
        Close()
    End Sub

    ' =========================================================
    ' MÉTODO FACTORY: Recopilar ítems del ensamblaje
    ' =========================================================

    Public Shared Function RecopilarItems(invApp As Inventor.Application, oAsmDoc As AssemblyDocument, sPrefix As String) As List(Of ComponenteRenombrarItem)
        Dim lista As New List(Of ComponenteRenombrarItem)()
        Dim visitados As New Dictionary(Of String, ComponenteRenombrarItem)(StringComparer.OrdinalIgnoreCase)

        ' 1. Incluir el ensamblaje raíz
        Try
            Dim sRootPath As String = oAsmDoc.FullFileName
            If Not String.IsNullOrEmpty(sRootPath) AndAlso IOFile.Exists(sRootPath) Then
                Dim sStock As String = GetProp(oAsmDoc, "Stock Number")
                Dim sPart As String = GetProp(oAsmDoc, "Part Number")
                Dim thumb As Image = ObtenerThumbnailSeguro(oAsmDoc)
                Dim sSugerido As String = If(Not String.IsNullOrEmpty(sPart), sPart, IOPath.GetFileNameWithoutExtension(sRootPath))

                Dim rootItem As New ComponenteRenombrarItem With {
                    .Occurrence = Nothing,
                    .Document = oAsmDoc,
                    .DocumentPath = sRootPath,
                    .FileName = IOPath.GetFileName(sRootPath),
                    .Extension = IOPath.GetExtension(sRootPath),
                    .IsAssembly = True,
                    .IsRootAssembly = True,
                    .StockNumber = sStock,
                    .PartNumber = sPart,
                    .SuggestedName = sSugerido,
                    .OriginalSuggestedName = sSugerido,
                    .Thumbnail = thumb
                }
                lista.Add(rootItem)
                visitados.Add(sRootPath, rootItem)
            End If
        Catch ex As Exception
        End Try

        ' 2. Recorrer ocurrencias recursivamente (soporta subensamblajes y piezas en cualquier profundidad)
        Dim allOccs As New List(Of ComponentOccurrence)()
        Try
            ObtenerOcurrenciasRecursivas(oAsmDoc.ComponentDefinition.Occurrences, allOccs)
        Catch ex As Exception
        End Try

        For Each occ As ComponentOccurrence In allOccs
            If occ Is Nothing Then Continue For

            Dim isSuppressed As Boolean = False
            Try
                isSuppressed = occ.Suppressed
            Catch
                isSuppressed = False
            End Try
            If isSuppressed Then Continue For

            Try
                Dim def As ComponentDefinition = occ.Definition
                If def IsNot Nothing Then
                    If def.BOMStructure = BOMStructureEnum.kPhantomBOMStructure OrElse def.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                        Continue For
                    End If
                End If
            Catch
            End Try

            Dim occDoc As Document = Nothing
            If Not TryGetDocument(occ, occDoc) OrElse occDoc Is Nothing Then
                Continue For
            End If

            Dim fullPath As String = ""
            Try
                fullPath = occDoc.FullFileName
            Catch
            End Try
            If String.IsNullOrEmpty(fullPath) OrElse Not IOFile.Exists(fullPath) Then Continue For

            Dim pathLower As String = fullPath.ToLowerInvariant()
            If pathLower.Contains("content center") OrElse pathLower.Contains("cc") OrElse pathLower.Contains("libraries") Then
                Continue For
            End If

            ' Descartar miembros derivados de estados de modelo si aplica
            If TypeOf occDoc Is PartDocument Then
                Try
                    Dim pDef = CType(occDoc, PartDocument).ComponentDefinition
                    If pDef.IsModelStateMember Then Continue For
                Catch
                End Try
            ElseIf TypeOf occDoc Is AssemblyDocument Then
                Try
                    Dim aDef = CType(occDoc, AssemblyDocument).ComponentDefinition
                    If aDef.IsModelStateMember Then Continue For
                Catch
                End Try
            End If

            ' Si ya se procesó este archivo único, asociar la ocurrencia si faltaba
            If visitados.ContainsKey(fullPath) Then
                If visitados(fullPath).Occurrence Is Nothing Then
                    visitados(fullPath).Occurrence = occ
                End If
                Continue For
            End If

            Dim sStockNum As String = GetProp(occDoc, "Stock Number")
            Dim sPartNum As String = GetProp(occDoc, "Part Number")
            Dim thumbImg As Image = ObtenerThumbnailSeguro(occDoc)
            Dim isAsm As Boolean = False
            Try
                isAsm = (occDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject)
            Catch
            End Try

            Dim sSugeridoDoc As String = If(Not String.IsNullOrEmpty(sPartNum), sPartNum, IOPath.GetFileNameWithoutExtension(fullPath))

            Dim item As New ComponenteRenombrarItem With {
                .Occurrence = occ,
                .Document = occDoc,
                .DocumentPath = fullPath,
                .FileName = IOPath.GetFileName(fullPath),
                .Extension = IOPath.GetExtension(fullPath),
                .IsAssembly = isAsm,
                .IsRootAssembly = False,
                .StockNumber = sStockNum,
                .PartNumber = sPartNum,
                .SuggestedName = sSugeridoDoc,
                .OriginalSuggestedName = sSugeridoDoc,
                .Thumbnail = thumbImg
            }

            lista.Add(item)
            visitados.Add(fullPath, item)
        Next

        Return lista
    End Function

    Private Shared Function TryGetDocument(occ As ComponentOccurrence, ByRef docOut As Document) As Boolean
        docOut = Nothing
        If occ Is Nothing Then Return False
        Try
            Dim def As ComponentDefinition = occ.Definition
            If def Is Nothing Then Return False

            If TypeOf def Is PartComponentDefinition Then
                Dim pDef = CType(def, PartComponentDefinition)
                If pDef.IsModelStateMember Then
                    docOut = CType(pDef.FactoryDocument, Document)
                Else
                    docOut = pDef.Document
                End If
                Return (docOut IsNot Nothing)
            ElseIf TypeOf def Is AssemblyComponentDefinition Then
                Dim aDef = CType(def, AssemblyComponentDefinition)
                If aDef.IsModelStateMember Then
                    docOut = CType(aDef.FactoryDocument, Document)
                Else
                    docOut = aDef.Document
                End If
                Return (docOut IsNot Nothing)
            Else
                Try
                    docOut = def.Document
                    Return (docOut IsNot Nothing)
                Catch
                    Return False
                End Try
            End If
        Catch
            Return False
        End Try
    End Function

    Private Shared Sub ObtenerOcurrenciasRecursivas(occs As Object, ByRef lista As List(Of ComponentOccurrence))
        If occs Is Nothing Then Return
        Dim enumerable As System.Collections.IEnumerable = TryCast(occs, System.Collections.IEnumerable)
        If enumerable Is Nothing Then Return

        For Each rawItem In enumerable
            Dim occ As ComponentOccurrence = TryCast(rawItem, ComponentOccurrence)
            If occ Is Nothing Then Continue For

            Dim isSuppressed As Boolean = False
            Try
                isSuppressed = occ.Suppressed
            Catch
                isSuppressed = False
            End Try
            If isSuppressed Then Continue For

            Try
                lista.Add(occ)

                Dim isAsm As Boolean = False
                Try
                    isAsm = (occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject)
                Catch
                    Try
                        isAsm = (occ.Definition.Document.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject)
                    Catch
                    End Try
                End Try

                If isAsm Then
                    Dim subOccs As Object = Nothing
                    Try
                        subOccs = occ.SubOccurrences
                    Catch
                        subOccs = Nothing
                    End Try

                    If subOccs IsNot Nothing Then
                        Dim count As Integer = 0
                        Try
                            count = CInt(CallByName(subOccs, "Count", CallType.Get))
                        Catch
                            count = 1
                        End Try

                        If count > 0 Then
                            ObtenerOcurrenciasRecursivas(subOccs, lista)
                        End If
                    End If
                End If
            Catch
            End Try
        Next
    End Sub

    Private Shared Function GetProp(oDoc As Document, propName As String) As String
        Dim val As String = ""
        Try : val = oDoc.PropertySets.Item("Design Tracking Properties").Item(propName).Value.ToString().Trim() : Catch : End Try
        If String.IsNullOrEmpty(val) Then
            Try : val = oDoc.PropertySets.Item("Inventor User Defined Properties").Item(propName).Value.ToString().Trim() : Catch : End Try
        End If
        Return val
    End Function

    ''' <summary>
    ''' Obtiene la miniatura (Thumbnail) de un documento de forma segura.
    ''' Se utiliza invocación dinámica por CallByName sobre Object para evitar
    ''' requerir referencia directa en tiempo de compilación al ensamblado 'stdole' (Error BC30652).
    ''' </summary>
    Private Shared Function ObtenerThumbnailSeguro(doc As Object) As Image
        Try
            If doc Is Nothing Then Return Nothing
            Dim pDisp As Object = CallByName(doc, "Thumbnail", CallType.Get)
            If pDisp IsNot Nothing Then
                Return PictureDispConverter.ToImage(pDisp)
            End If
        Catch
        End Try
        Return Nothing
    End Function

End Class
