Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports Inventor

Imports IOPath = System.IO.Path
Imports IOFile = System.IO.File

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
            Dim tipoStr As String = If(IsRootAssembly, "[Raíz]", If(IsAssembly, "[Ensamblaje]", "[Pieza]"))
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

    Private _invApp As Inventor.Application
    Private _items As List(Of ComponenteRenombrarItem)
    Private _prefix As String

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
            .BackColor = System.Drawing.Color.FromArgb(248, 250, 252)
        }

        Dim lblTitulo As New Label With {
            .Text = "Revisión y Asignación de Nombres a Componentes",
            .Font = New Font("Segoe UI", 10.0F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(15, 23, 42),
            .Location = New System.Drawing.Point(12, 8),
            .AutoSize = True
        }

        Dim lblDesc As New Label With {
            .Text = "Verifique los nombres sugeridos basados en Part Number. Puede editar el campo 'Nombre Sugerido' directamente en la celda." & vbCrLf & "Haga doble clic en una fila o use los botones de la derecha para hacer zoom o abrir la pieza.",
            .ForeColor = System.Drawing.Color.FromArgb(100, 116, 139),
            .Location = New System.Drawing.Point(12, 28),
            .AutoSize = True
        }

        _btnRestaurar = New Button With {
            .Text = "🔄 Restaurar Sugeridos Originales",
            .Location = New System.Drawing.Point(680, 10),
            .Size = New Size(220, 26),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .BackColor = System.Drawing.Color.White,
            .Font = New Font("Segoe UI", 8.5F)
        }
        AddHandler _btnRestaurar.Click, AddressOf OnRestaurarSugeridos

        Dim lblCmb As New Label With {
            .Text = "Buscar / Seleccionar:",
            .Location = New System.Drawing.Point(530, 43),
            .AutoSize = True,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        }

        _cmbPiezas = New ComboBox With {
            .Location = New System.Drawing.Point(680, 40),
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
            .BackColor = System.Drawing.Color.FromArgb(241, 245, 249)
        }

        Dim grpPreview As New GroupBox With {
            .Text = "Detalle del Componente",
            .Dock = DockStyle.Fill,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(30, 41, 59)
        }

        _picPreview = New PictureBox With {
            .Location = New System.Drawing.Point(15, 22),
            .Size = New Size(210, 160),
            .SizeMode = PictureBoxSizeMode.Zoom,
            .BorderStyle = BorderStyle.FixedSingle,
            .BackColor = System.Drawing.Color.White
        }

        _lblTipo = New Label With {
            .Location = New System.Drawing.Point(15, 190),
            .Size = New Size(210, 18),
            .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(71, 85, 105),
            .Text = "Tipo: -"
        }

        _lblStockNumber = New Label With {
            .Location = New System.Drawing.Point(15, 212),
            .Size = New Size(210, 32),
            .Font = New Font("Segoe UI", 8.0F),
            .Text = "Stock Number: -"
        }

        _lblPartNumber = New Label With {
            .Location = New System.Drawing.Point(15, 246),
            .Size = New Size(210, 32),
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(2, 132, 199),
            .Text = "Part Number: -"
        }

        _lblActual = New Label With {
            .Location = New System.Drawing.Point(15, 280),
            .Size = New Size(210, 32),
            .Font = New Font("Segoe UI", 8.0F),
            .ForeColor = System.Drawing.Color.FromArgb(100, 116, 139),
            .Text = "Archivo actual: -"
        }

        _lblNuevo = New Label With {
            .Location = New System.Drawing.Point(15, 314),
            .Size = New Size(210, 36),
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(16, 185, 129),
            .Text = "Nuevo archivo: -"
        }

        _btnZoom = New Button With {
            .Text = "🔍 Hacer Zoom",
            .Location = New System.Drawing.Point(15, 358),
            .Size = New Size(210, 32),
            .BackColor = System.Drawing.Color.White,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(30, 41, 59),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler _btnZoom.Click, AddressOf OnBtnZoomClick

        _btnAbrir = New Button With {
            .Text = "📂 Abrir Elemento",
            .Location = New System.Drawing.Point(15, 396),
            .Size = New Size(210, 32),
            .BackColor = System.Drawing.Color.White,
            .Font = New Font("Segoe UI", 9.0F),
            .ForeColor = System.Drawing.Color.FromArgb(30, 41, 59),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler _btnAbrir.Click, AddressOf OnBtnAbrirClick

        Dim lblTip As New Label With {
            .Text = "💡 Tip: Doble clic en una fila para hacer zoom directamente.",
            .Location = New System.Drawing.Point(15, 436),
            .Size = New Size(210, 36),
            .Font = New Font("Segoe UI", 7.5F, FontStyle.Italic),
            .ForeColor = System.Drawing.Color.FromArgb(148, 163, 184)
        }

        ' Botones de acción directa en el panel lateral derecho
        _btnAplicarLateral = New Button With {
            .Text = "✔️ Aplicar y Renombrar",
            .Location = New System.Drawing.Point(15, 478),
            .Size = New Size(210, 34),
            .BackColor = System.Drawing.Color.FromArgb(2, 132, 199),
            .ForeColor = System.Drawing.Color.White,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        _btnAplicarLateral.FlatAppearance.BorderSize = 0
        AddHandler _btnAplicarLateral.Click, AddressOf OnAplicar

        _btnCancelarLateral = New Button With {
            .Text = "✖️ Cancelar",
            .Location = New System.Drawing.Point(15, 518),
            .Size = New Size(210, 30),
            .BackColor = System.Drawing.Color.White,
            .ForeColor = System.Drawing.Color.FromArgb(51, 65, 85),
            .Font = New Font("Segoe UI", 8.5F),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        _btnCancelarLateral.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(203, 213, 225)
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
            .BackColor = System.Drawing.Color.FromArgb(241, 245, 249)
        }

        Dim sepBottom As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 1,
            .BackColor = System.Drawing.Color.FromArgb(203, 213, 225)
        }
        pnlBottom.Controls.Add(sepBottom)

        _lblResumen = New Label With {
            .Location = New System.Drawing.Point(16, 18),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(51, 65, 85),
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
            .BackColor = System.Drawing.Color.White,
            .ForeColor = System.Drawing.Color.FromArgb(51, 65, 85),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand,
            .Margin = New Padding(10, 0, 0, 0)
        }
        _btnCancelar.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(203, 213, 225)
        AddHandler _btnCancelar.Click, AddressOf OnCancelar

        _btnAplicar = New Button With {
            .Text = "Aplicar y Renombrar",
            .Size = New Size(170, 34),
            .BackColor = System.Drawing.Color.FromArgb(2, 132, 199),
            .ForeColor = System.Drawing.Color.White,
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
            .BackgroundColor = System.Drawing.Color.White,
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
                .BackColor = System.Drawing.Color.FromArgb(254, 252, 232),
                .SelectionBackColor = System.Drawing.Color.FromArgb(253, 224, 71),
                .SelectionForeColor = System.Drawing.Color.Black
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
            _lblNuevo.ForeColor = System.Drawing.Color.FromArgb(16, 185, 129)
        Else
            _lblNuevo.ForeColor = System.Drawing.Color.FromArgb(100, 116, 139)
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
    ' ACCIONES: ZOOM Y ABRIR ELEMENTO
    ' =========================================================

    Public Sub HacerZoom(item As ComponenteRenombrarItem)
        If item Is Nothing OrElse _invApp Is Nothing Then Return
        Try
            Dim doc As Document = _invApp.ActiveDocument
            If doc Is Nothing Then Return

            If item.IsRootAssembly OrElse item.Occurrence Is Nothing Then
                If _invApp.ActiveView IsNot Nothing Then
                    _invApp.ActiveView.Fit()
                    _invApp.ActiveView.Update()
                End If
                Return
            End If

            Dim occ As ComponentOccurrence = item.Occurrence
            Try
                doc.SelectSet.Clear()
                doc.SelectSet.Select(occ)
            Catch
            End Try

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
        If item Is Nothing OrElse String.IsNullOrEmpty(item.DocumentPath) OrElse _invApp Is Nothing Then Return
        Try
            If Not IOFile.Exists(item.DocumentPath) Then
                MessageBox.Show("El archivo no existe en disco:" & vbCrLf & item.DocumentPath, "Archivo no encontrado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim docExistente As Document = Nothing
            For Each d As Document In _invApp.Documents
                If String.Equals(d.FullFileName, item.DocumentPath, StringComparison.OrdinalIgnoreCase) Then
                    docExistente = d
                    Exit For
                End If
            Next

            If docExistente IsNot Nothing Then
                docExistente.Activate()
            Else
                _invApp.Documents.Open(item.DocumentPath, True)
            End If
        Catch ex As Exception
            MessageBox.Show("No se pudo abrir el elemento: " & ex.Message, "Error al abrir", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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

        ' 2. Recorrer ocurrencias recursivamente
        Dim allOccs As New List(Of ComponentOccurrence)()
        ObtenerOcurrenciasRecursivas(oAsmDoc.ComponentDefinition.Occurrences, allOccs)

        For Each occ As ComponentOccurrence In allOccs
            If occ Is Nothing Then Continue For
            If occ.Suppressed Then Continue For

            Try
                Dim def As ComponentDefinition = occ.Definition
                If def.BOMStructure = BOMStructureEnum.kPhantomBOMStructure OrElse def.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                    Continue For
                End If
            Catch
            End Try

            Dim occDoc As Document = Nothing
            Try
                occDoc = occ.Definition.Document
            Catch
                Continue For
            End Try

            If occDoc Is Nothing Then Continue For
            Dim fullPath As String = occDoc.FullFileName
            If String.IsNullOrEmpty(fullPath) OrElse Not IOFile.Exists(fullPath) Then Continue For

            Dim pathLower As String = fullPath.ToLowerInvariant()
            If pathLower.Contains("content center") OrElse pathLower.Contains("cc") OrElse pathLower.Contains("libraries") Then
                Continue For
            End If

            If TypeOf occDoc Is PartDocument Then
                Try
                    Dim pDef = CType(occDoc, PartDocument).ComponentDefinition
                    If pDef.IsModelStateMember Then Continue For
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
            Dim isAsm As Boolean = (occDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject)
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

    Private Shared Sub ObtenerOcurrenciasRecursivas(occs As ComponentOccurrences, ByRef lista As List(Of ComponentOccurrence))
        For Each occ As ComponentOccurrence In occs
            lista.Add(occ)
            If occ.SubOccurrences IsNot Nothing AndAlso occ.SubOccurrences.Count > 0 Then
                ObtenerOcurrenciasRecursivas(occ.SubOccurrences, lista)
            End If
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
