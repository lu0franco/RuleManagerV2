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
''' Item que representa una pieza con propiedad MEC / MC para la revisión de sobremedida de largo.
''' </summary>
Public Class ElementoMecItem
    Public Property Occurrence As ComponentOccurrence
    Public Property Document As Document
    Public Property DocumentPath As String
    Public Property FileName As String
    Public Property PartNumber As String
    Public Property StockNumber As String
    Public Property MaterialPlano As String
    Public Property LargoOriginalStr As String
    Public Property LargoOriginalNum As Double
    Public Property AnchoStr As String
    Public Property Tipo As String = "MEC" ' "MEC" o "MC"
    Public Property Sumar3mm As Boolean = True
    Public Property Thumbnail As Image
    Public Property PropLargo As Inventor.Property

    Public ReadOnly Property ColumnaPiezaStock As String
        Get
            Dim tipoStr As String = "[" & Tipo & "]"
            If Not String.IsNullOrEmpty(StockNumber) Then
                Return StockNumber & " " & tipoStr
            Else
                Return If(Not String.IsNullOrEmpty(PartNumber), PartNumber, FileName) & " " & tipoStr
            End If
        End Get
    End Property

    Public ReadOnly Property LargoFinalStr As String
        Get
            If Sumar3mm Then
                Dim valFinal As Double = LargoOriginalNum + 3.0
                Return valFinal.ToString() & " mm"
            Else
                Return LargoOriginalStr
            End If
        End Get
    End Property

    Public ReadOnly Property TieneCambio As Boolean
        Get
            Return Sumar3mm
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return If(Not String.IsNullOrEmpty(PartNumber), PartNumber, FileName) & " (" & If(Not String.IsNullOrEmpty(StockNumber), StockNumber, Tipo) & ")"
    End Function
End Class

''' <summary>
''' Ventana interactiva que muestra todos los elementos con MC/MEC = '✓',
''' permitiendo elegir con la columna SUMA si agregar +3mm de largo,
''' con selector desplegable, visor de vista previa 3D, botones de zoom,
''' abrir elemento, marcar/desmarcar todos, cancelar y aplicar cambios.
''' </summary>
Public Class SobremedidaLargoDialog
    Inherits Form

    <DllImport("user32.dll")>
    Private Shared Function EnableWindow(hWnd As IntPtr, bEnable As Boolean) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function

    Private _invApp As Inventor.Application
    Private _items As List(Of ElementoMecItem)
    Private _highlightSet As HighlightSet = Nothing

    Private _grid As DataGridView
    Private _cmbPiezas As ComboBox
    Private _picPreview As PictureBox
    Private _lblTipo As Label
    Private _lblPartNumber As Label
    Private _lblStockNumber As Label
    Private _lblMaterial As Label
    Private _lblLargoOriginal As Label
    Private _lblLargoFinal As Label
    Private _lblResumen As Label
    Private _btnMarcarTodos As Button
    Private _btnDesmarcarTodos As Button
    Private _btnZoom As Button
    Private _btnAbrir As Button
    Private _btnAceptar As Button
    Private _btnCancelar As Button
    Private _btnAceptarLateral As Button
    Private _btnCancelarLateral As Button
    Private _contextMenu As ContextMenuStrip

    Public ReadOnly Property Items As List(Of ElementoMecItem)
        Get
            Return _items
        End Get
    End Property

    Public Sub New(invApp As Inventor.Application, items As List(Of ElementoMecItem))
        _invApp = invApp
        _items = If(items, New List(Of ElementoMecItem)())
        InitializeComponent()
        CargarDatos()
    End Sub

    Public Sub New(items As List(Of ElementoMecItem))
        _items = If(items, New List(Of ElementoMecItem)())
        Try
            If _items.Count > 0 AndAlso _items(0).Occurrence IsNot Nothing Then
                _invApp = TryCast(_items(0).Occurrence.Application, Inventor.Application)
            End If
        Catch
        End Try
        InitializeComponent()
        CargarDatos()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Sobremedida de Largo (+3mm) - Revisión General de Componentes MEC / MC"
        Me.Size = New Size(1020, 640)
        Me.MinimumSize = New Size(850, 520)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Font = New Font("Segoe UI", 9.0F)

        ' =========================================================
        ' PANEL SUPERIOR: Título, instrucciones y acciones rápidas
        ' =========================================================
        Dim pnlTop As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 72,
            .Padding = New Padding(12, 8, 12, 8),
            .BackColor = System.Drawing.Color.FromArgb(248, 250, 252)
        }

        Dim lblTitulo As New Label With {
            .Text = "Revisión de Largo para Componentes con marca MEC / MC (✓)",
            .Font = New Font("Segoe UI", 10.0F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(15, 23, 42),
            .Location = New System.Drawing.Point(12, 8),
            .AutoSize = True
        }

        Dim lblDesc As New Label With {
            .Text = "Active o desactive la casilla 'SUMA (+3mm)' para definir qué piezas llevarán sobremedida." & vbCrLf & "Haga doble clic en una fila o use los botones de la derecha para hacer zoom 3D o abrir la pieza.",
            .ForeColor = System.Drawing.Color.FromArgb(100, 116, 139),
            .Location = New System.Drawing.Point(12, 28),
            .AutoSize = True
        }

        _btnMarcarTodos = New Button With {
            .Text = "☑ Marcar Todos (+3mm)",
            .Location = New System.Drawing.Point(680, 10),
            .Size = New Size(155, 26),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .BackColor = System.Drawing.Color.White,
            .Font = New Font("Segoe UI", 8.5F)
        }
        AddHandler _btnMarcarTodos.Click, AddressOf OnMarcarTodos

        _btnDesmarcarTodos = New Button With {
            .Text = "☐ Desmarcar Todos",
            .Location = New System.Drawing.Point(842, 10),
            .Size = New Size(148, 26),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .BackColor = System.Drawing.Color.White,
            .Font = New Font("Segoe UI", 8.5F)
        }
        AddHandler _btnDesmarcarTodos.Click, AddressOf OnDesmarcarTodos

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
        pnlTop.Controls.Add(_btnMarcarTodos)
        pnlTop.Controls.Add(_btnDesmarcarTodos)
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

        _lblPartNumber = New Label With {
            .Location = New System.Drawing.Point(15, 212),
            .Size = New Size(210, 32),
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(2, 132, 199),
            .Text = "Part Number: -"
        }

        _lblStockNumber = New Label With {
            .Location = New System.Drawing.Point(15, 246),
            .Size = New Size(210, 26),
            .Font = New Font("Segoe UI", 8.0F),
            .Text = "Stock Number: -"
        }

        _lblMaterial = New Label With {
            .Location = New System.Drawing.Point(15, 274),
            .Size = New Size(210, 26),
            .Font = New Font("Segoe UI", 8.0F),
            .ForeColor = System.Drawing.Color.FromArgb(100, 116, 139),
            .Text = "Material Plano: -"
        }

        _lblLargoOriginal = New Label With {
            .Location = New System.Drawing.Point(15, 302),
            .Size = New Size(210, 20),
            .Font = New Font("Segoe UI", 8.0F),
            .ForeColor = System.Drawing.Color.FromArgb(100, 116, 139),
            .Text = "Largo Nominal: -"
        }

        _lblLargoFinal = New Label With {
            .Location = New System.Drawing.Point(15, 324),
            .Size = New Size(210, 28),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(16, 185, 129),
            .Text = "Largo Final: -"
        }

        _btnZoom = New Button With {
            .Text = "🔍 Hacer Zoom",
            .Location = New System.Drawing.Point(15, 358),
            .Size = New Size(210, 32),
            .BackColor = System.Drawing.Color.White,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(30, 41, 59),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        AddHandler _btnZoom.Click, AddressOf OnBtnZoomClick

        _btnAbrir = New Button With {
            .Text = "📂 Abrir Elemento",
            .Location = New System.Drawing.Point(15, 396),
            .Size = New Size(210, 32),
            .BackColor = System.Drawing.Color.White,
            .Font = New Font("Segoe UI", 9.0F),
            .ForeColor = System.Drawing.Color.FromArgb(30, 41, 59),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
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
        _btnAceptarLateral = New Button With {
            .Text = "✔️ Aceptar y Aplicar (+3mm)",
            .Location = New System.Drawing.Point(15, 478),
            .Size = New Size(210, 34),
            .BackColor = System.Drawing.Color.FromArgb(2, 132, 199),
            .ForeColor = System.Drawing.Color.White,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        _btnAceptarLateral.FlatAppearance.BorderSize = 0
        AddHandler _btnAceptarLateral.Click, AddressOf OnAceptar

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
        grpPreview.Controls.Add(_lblPartNumber)
        grpPreview.Controls.Add(_lblStockNumber)
        grpPreview.Controls.Add(_lblMaterial)
        grpPreview.Controls.Add(_lblLargoOriginal)
        grpPreview.Controls.Add(_lblLargoFinal)
        grpPreview.Controls.Add(_btnZoom)
        grpPreview.Controls.Add(_btnAbrir)
        grpPreview.Controls.Add(lblTip)
        grpPreview.Controls.Add(_btnAceptarLateral)
        grpPreview.Controls.Add(_btnCancelarLateral)
        pnlRight.Controls.Add(grpPreview)

        ' =========================================================
        ' PANEL INFERIOR: Resumen y Botones Aceptar / Cancelar
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

        _btnAceptar = New Button With {
            .Text = "Aceptar y Aplicar (+3mm)",
            .Size = New Size(195, 34),
            .BackColor = System.Drawing.Color.FromArgb(2, 132, 199),
            .ForeColor = System.Drawing.Color.White,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand,
            .Margin = New Padding(0, 0, 0, 0)
        }
        _btnAceptar.FlatAppearance.BorderSize = 0
        AddHandler _btnAceptar.Click, AddressOf OnAceptar

        pnlBotonesInferiores.Controls.Add(_btnCancelar)
        pnlBotonesInferiores.Controls.Add(_btnAceptar)

        pnlBottom.Controls.Add(pnlBotonesInferiores)
        pnlBottom.Controls.Add(_lblResumen)

        AddHandler pnlBottom.Resize, Sub(s, e)
                                         _lblResumen.Top = (pnlBottom.ClientSize.Height - _lblResumen.Height) \ 2
                                     End Sub

        ' =========================================================
        ' CENTRO: DataGridView con las columnas
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

        Me.AcceptButton = _btnAceptar
        Me.CancelButton = _btnCancelar

        AddHandler Me.Shown, Sub(s, e) HabilitarVentanaInventor()
        AddHandler Me.Activated, Sub(s, e) HabilitarVentanaInventor()
        AddHandler Me.FormClosing, Sub(s, e)
                                       LimpiarResaltado()
                                       HabilitarVentanaInventor()
                                   End Sub
    End Sub

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
    ''' al HighlightSet para asegurar que toda la geometría se coloree en el visor 3D.
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
            If occ.SubOccurrences IsNot Nothing AndAlso occ.SubOccurrences.Count > 0 Then
                For Each subOcc As ComponentOccurrence In occ.SubOccurrences
                    If Not subOcc.Suppressed Then
                        AgregarOcurrenciaAHighlight(subOcc, hs)
                    End If
                Next
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

                ' Resaltar la ocurrencia y sus cuerpos sólidos 3D
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

    Private Sub ConfigurarColumnasGrid()
        ' Columna 1: SUMA (+3mm) CheckBox
        Dim colSuma As New DataGridViewCheckBoxColumn()
        colSuma.HeaderText = "SUMA (+3mm)"
        colSuma.Name = "colSuma"
        colSuma.Width = 95
        colSuma.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        colSuma.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        colSuma.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(248, 250, 252)

        ' Columna 2: Part Number
        Dim colPart As New DataGridViewTextBoxColumn With {
            .HeaderText = "Part Number",
            .Name = "colPart",
            .ReadOnly = True,
            .Width = 150
        }

        ' Columna 3: Stock Number
        Dim colStock As New DataGridViewTextBoxColumn With {
            .HeaderText = "Stock Number",
            .Name = "colStock",
            .ReadOnly = True,
            .Width = 130
        }

        ' Columna 4: Material Plano
        Dim colMat As New DataGridViewTextBoxColumn With {
            .HeaderText = "Material Plano",
            .Name = "colMat",
            .ReadOnly = True,
            .Width = 140
        }

        ' Columna 5: Largo Original
        Dim colLargoOrig As New DataGridViewTextBoxColumn With {
            .HeaderText = "Largo Nom.",
            .Name = "colLargoOrig",
            .ReadOnly = True,
            .Width = 95,
            .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight}
        }

        ' Columna 6: Largo Final (+3mm)
        Dim colLargoFinal As New DataGridViewTextBoxColumn With {
            .HeaderText = "Largo Final",
            .Name = "colLargoFinal",
            .ReadOnly = True,
            .Width = 110,
            .DefaultCellStyle = New DataGridViewCellStyle With {
                .Alignment = DataGridViewContentAlignment.MiddleRight,
                .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
            }
        }

        ' Columna 7: Ancho
        Dim colAncho As New DataGridViewTextBoxColumn With {
            .HeaderText = "Ancho",
            .Name = "colAncho",
            .ReadOnly = True,
            .Width = 85,
            .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight}
        }

        ' Columna 8: Tipo (MEC / MC)
        Dim colTipo As New DataGridViewTextBoxColumn With {
            .HeaderText = "Tipo",
            .Name = "colTipo",
            .ReadOnly = True,
            .Width = 65,
            .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleCenter}
        }

        _grid.Columns.AddRange(New DataGridViewColumn() {
            colSuma, colPart, colStock, colMat, colLargoOrig, colLargoFinal, colAncho, colTipo
        })
    End Sub

    Private Sub ConfigurarMenuContextual()
        _contextMenu = New ContextMenuStrip()

        Dim mnuZoom As New ToolStripMenuItem("🔍 Hacer Zoom en 3D")
        AddHandler mnuZoom.Click, AddressOf OnBtnZoomClick
        _contextMenu.Items.Add(mnuZoom)

        Dim mnuAbrir As New ToolStripMenuItem("📂 Abrir Archivo en Inventor")
        AddHandler mnuAbrir.Click, AddressOf OnBtnAbrirClick
        _contextMenu.Items.Add(mnuAbrir)

        _contextMenu.Items.Add(New ToolStripSeparator())

        Dim mnuSumar As New ToolStripMenuItem("☑️ Sumar +3mm a esta pieza")
        AddHandler mnuSumar.Click, Sub()
                                       Dim it = GetSelectedItem()
                                       If it IsNot Nothing AndAlso _grid.CurrentRow IsNot Nothing Then
                                           it.Sumar3mm = True
                                           _grid.CurrentRow.Cells("colSuma").Value = True
                                           ActualizarFilaLargo(_grid.CurrentRow.Index)
                                           ActualizarResumen()
                                       End If
                                   End Sub
        _contextMenu.Items.Add(mnuSumar)

        Dim mnuQuitar As New ToolStripMenuItem("☐ Quitar +3mm (Dejar nominal)")
        AddHandler mnuQuitar.Click, Sub()
                                        Dim it = GetSelectedItem()
                                        If it IsNot Nothing AndAlso _grid.CurrentRow IsNot Nothing Then
                                            it.Sumar3mm = False
                                            _grid.CurrentRow.Cells("colSuma").Value = False
                                            ActualizarFilaLargo(_grid.CurrentRow.Index)
                                            ActualizarResumen()
                                        End If
                                    End Sub
        _contextMenu.Items.Add(mnuQuitar)

        _contextMenu.Items.Add(New ToolStripSeparator())

        Dim mnuMarcarTodos As New ToolStripMenuItem("☑️ Marcar Todos (+3mm)")
        AddHandler mnuMarcarTodos.Click, AddressOf OnMarcarTodos
        _contextMenu.Items.Add(mnuMarcarTodos)

        Dim mnuDesmarcarTodos As New ToolStripMenuItem("☐ Desmarcar Todos")
        AddHandler mnuDesmarcarTodos.Click, AddressOf OnDesmarcarTodos
        _contextMenu.Items.Add(mnuDesmarcarTodos)

        _grid.ContextMenuStrip = _contextMenu
    End Sub

    Private Sub CargarDatos()
        _grid.Rows.Clear()
        _cmbPiezas.Items.Clear()

        For Each item In _items
            ' Cargar thumbnail de forma segura si no fue precargado
            If item.Thumbnail Is Nothing AndAlso item.Occurrence IsNot Nothing Then
                Try
                    Dim oDoc As Document = item.Occurrence.Definition.Document
                    If oDoc IsNot Nothing Then
                        item.Thumbnail = ObtenerThumbnailSeguro(oDoc)
                    End If
                Catch
                End Try
            End If

            Dim rowIdx = _grid.Rows.Add(
                item.Sumar3mm,
                item.PartNumber,
                item.StockNumber,
                item.MaterialPlano,
                item.LargoOriginalStr,
                item.LargoFinalStr,
                item.AnchoStr,
                item.Tipo
            )
            _grid.Rows(rowIdx).Tag = item

            ' Color inicial según si tiene suma
            If item.Sumar3mm Then
                _grid.Rows(rowIdx).Cells("colLargoFinal").Style.ForeColor = System.Drawing.Color.FromArgb(16, 185, 129)
            Else
                _grid.Rows(rowIdx).Cells("colLargoFinal").Style.ForeColor = System.Drawing.Color.FromArgb(100, 116, 139)
            End If

            _cmbPiezas.Items.Add(item)
        Next

        If _cmbPiezas.Items.Count > 0 Then
            _cmbPiezas.SelectedIndex = 0
        End If

        ActualizarResumen()
    End Sub

    Private Function GetSelectedItem() As ElementoMecItem
        If _grid.CurrentRow IsNot Nothing AndAlso _grid.CurrentRow.Tag IsNot Nothing Then
            Return DirectCast(_grid.CurrentRow.Tag, ElementoMecItem)
        End If
        If _cmbPiezas.SelectedItem IsNot Nothing Then
            Return DirectCast(_cmbPiezas.SelectedItem, ElementoMecItem)
        End If
        Return Nothing
    End Function

    Private Sub ActualizarFilaLargo(rowIdx As Integer)
        If rowIdx < 0 OrElse rowIdx >= _grid.Rows.Count Then Exit Sub
        Dim item = DirectCast(_grid.Rows(rowIdx).Tag, ElementoMecItem)
        If item Is Nothing Then Exit Sub

        _grid.Rows(rowIdx).Cells("colLargoFinal").Value = item.LargoFinalStr

        If item.Sumar3mm Then
            _grid.Rows(rowIdx).Cells("colLargoFinal").Style.ForeColor = System.Drawing.Color.FromArgb(16, 185, 129)
        Else
            _grid.Rows(rowIdx).Cells("colLargoFinal").Style.ForeColor = System.Drawing.Color.FromArgb(100, 116, 139)
        End If

        If _grid.CurrentRow IsNot Nothing AndAlso _grid.CurrentRow.Index = rowIdx Then
            ActualizarDetalleItem(item)
        End If
    End Sub

    Private Sub ActualizarResumen()
        Dim total = _items.Count
        Dim conSuma = 0
        For Each it In _items
            If it.Sumar3mm Then conSuma += 1
        Next
        Dim sinSuma = total - conSuma
        _lblResumen.Text = "Total componentes: " & total & "  |  Con sobremedida (+3mm): " & conSuma & "  |  Sin sobremedida: " & sinSuma
    End Sub

    Private Sub ActualizarDetalleItem(item As ElementoMecItem)
        If item Is Nothing Then
            _picPreview.Image = Nothing
            _lblTipo.Text = "Tipo: -"
            _lblPartNumber.Text = "Part Number: -"
            _lblStockNumber.Text = "Stock Number: -"
            _lblMaterial.Text = "Material Plano: -"
            _lblLargoOriginal.Text = "Largo Nominal: -"
            _lblLargoFinal.Text = "Largo Final: -"
            Return
        End If

        _picPreview.Image = item.Thumbnail
        _lblTipo.Text = "Tipo: " & If(String.IsNullOrEmpty(item.Tipo), "Pieza MEC", item.Tipo)
        _lblPartNumber.Text = "Part Number: " & item.PartNumber
        _lblStockNumber.Text = "Stock Number: " & If(String.IsNullOrEmpty(item.StockNumber), "(Sin Stock)", item.StockNumber)
        _lblMaterial.Text = "Material Plano: " & If(String.IsNullOrEmpty(item.MaterialPlano), "-", item.MaterialPlano)
        _lblLargoOriginal.Text = "Largo Nominal: " & item.LargoOriginalStr

        If item.Sumar3mm Then
            _lblLargoFinal.ForeColor = System.Drawing.Color.FromArgb(16, 185, 129)
            _lblLargoFinal.Text = "Largo Final: " & item.LargoFinalStr & " (+3mm)"
        Else
            _lblLargoFinal.ForeColor = System.Drawing.Color.FromArgb(100, 116, 139)
            _lblLargoFinal.Text = "Largo Final: " & item.LargoFinalStr & " (Nominal)"
        End If
    End Sub

    ' =========================================================
    ' ACCIONES: ZOOM Y ABRIR ELEMENTO
    ' =========================================================

    Public Sub HacerZoom(item As ElementoMecItem)
        If item Is Nothing OrElse _invApp Is Nothing Then Return
        Try
            If item.Occurrence Is Nothing Then
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

    Public Sub AbrirElemento(item As ElementoMecItem)
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
        If e.ColumnIndex <> _grid.Columns("colSuma").Index Then
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
        If e.ColumnIndex = _grid.Columns("colSuma").Index Then
            Dim item = DirectCast(_grid.Rows(e.RowIndex).Tag, ElementoMecItem)
            If item IsNot Nothing Then
                Dim val = _grid.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                item.Sumar3mm = If(val IsNot Nothing AndAlso TypeOf val Is Boolean, CBool(val), False)
                ActualizarFilaLargo(e.RowIndex)
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
        End If
    End Sub

    Private Sub OnCmbSelectedIndexChanged(sender As Object, e As EventArgs)
        Dim item = DirectCast(_cmbPiezas.SelectedItem, ElementoMecItem)
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

            If item.Occurrence IsNot Nothing Then
                ResaltarElemento(item.Occurrence)
            End If
        End If
    End Sub

    Private Sub OnMarcarTodos(sender As Object, e As EventArgs)
        For Each row As DataGridViewRow In _grid.Rows
            Dim item = DirectCast(row.Tag, ElementoMecItem)
            If item IsNot Nothing Then
                item.Sumar3mm = True
                row.Cells("colSuma").Value = True
                ActualizarFilaLargo(row.Index)
            End If
        Next
        ActualizarResumen()
    End Sub

    Private Sub OnDesmarcarTodos(sender As Object, e As EventArgs)
        For Each row As DataGridViewRow In _grid.Rows
            Dim item = DirectCast(row.Tag, ElementoMecItem)
            If item IsNot Nothing Then
                item.Sumar3mm = False
                row.Cells("colSuma").Value = False
                ActualizarFilaLargo(row.Index)
            End If
        Next
        ActualizarResumen()
    End Sub

    ''' <summary>
    ''' Aplica los valores de Largo final en la propiedad LARGO de cada componente.
    ''' </summary>
    Public Sub AplicarCambios()
        For Each item In _items
            Try
                If item.PropLargo IsNot Nothing Then
                    item.PropLargo.Value = item.LargoFinalStr
                End If
            Catch
            End Try
        Next
    End Sub

    Private Sub OnAceptar(sender As Object, e As EventArgs)
        AplicarCambios()
        LimpiarResaltado()
        MessageBox.Show("Se han aplicado los cambios de sobremedida (+3mm) correctamente." & vbCrLf & vbCrLf &
                        "• Piezas MEC revisadas: " & _items.Count,
                        "Asignar Propiedades", MessageBoxButtons.OK, MessageBoxIcon.Information)
        DialogResult = DialogResult.OK
        Close()
    End Sub

    Private Sub OnCancelar(sender As Object, e As EventArgs)
        LimpiarResaltado()
        DialogResult = DialogResult.Cancel
        Close()
    End Sub

    ''' <summary>
    ''' Método estático que ejecuta la revisión interactiva mostrando el diálogo de forma no modal.
    ''' </summary>
    Public Shared Function EjecutarRevision(invApp As Inventor.Application, items As List(Of ElementoMecItem)) As Boolean
        If items Is Nothing OrElse items.Count = 0 Then
            Return False
        End If

        Dim hwnd As IntPtr = IntPtr.Zero
        Try
            If invApp IsNot Nothing Then hwnd = New IntPtr(invApp.MainFrameHWND)
        Catch
        End Try

        Dim dlg As New SobremedidaLargoDialog(invApp, items)
        If hwnd <> IntPtr.Zero Then
            dlg.Show(New WindowWrapper(hwnd))
        Else
            dlg.Show()
        End If
        Return True
    End Function

    ''' <summary>
    ''' Extrae la miniatura del documento de manera segura usando late-binding reflection,
    ''' evitando la necesidad de referenciar stdole.dll / IPictureDisp en tiempo de compilación.
    ''' </summary>
    Private Shared Function ObtenerThumbnailSeguro(oDoc As Document) As Image
        If oDoc Is Nothing Then Return Nothing
        Try
            Dim pDisp As Object = CallByName(oDoc, "Thumbnail", CallType.Get)
            If pDisp IsNot Nothing Then
                Return PictureDispConverter.ToImage(pDisp)
            End If
        Catch
        End Try
        Return Nothing
    End Function

End Class
