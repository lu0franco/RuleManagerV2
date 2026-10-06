Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Windows.Forms
Imports Inventor

''' <summary>
''' Item que representa una pieza con propiedad MEC / MC para la revisión de sobremedida de largo.
''' </summary>
Public Class ElementoMecItem
    Public Property Occurrence As ComponentOccurrence
    Public Property DocumentPath As String
    Public Property PartNumber As String
    Public Property StockNumber As String
    Public Property MaterialPlano As String
    Public Property LargoOriginalStr As String
    Public Property LargoOriginalNum As Double
    Public Property AnchoStr As String
    Public Property Tipo As String ' "MEC" o "MC"
    Public Property Sumar3mm As Boolean = True
    Public Property Thumbnail As Image
    Public Property PropLargo As Inventor.Property

    Public ReadOnly Property LargoFinalStr As String
        Get
            If Sumar3mm Then
                Return (LargoOriginalNum + 3.0).ToString() & " mm"
            Else
                Return LargoOriginalStr
            End If
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return PartNumber & " (" & StockNumber & ") [" & Tipo & "]"
    End Function
End Class

''' <summary>
''' Ventana interactiva que muestra todos los elementos con MC/MEC = '✓',
''' permitiendo elegir con la columna SUMA si agregar +3mm de largo,
''' junto con selector desplegable y visor de vista previa 3D de la pieza.
''' </summary>
Public Class SobremedidaLargoDialog
    Inherits Form

    Private _items As List(Of ElementoMecItem)
    Private _grid As DataGridView
    Private _cmbPiezas As ComboBox
    Private _picPreview As PictureBox
    Private _lblPartNumber As Label
    Private _lblStockNumber As Label
    Private _lblMaterial As Label
    Private _lblLargoFinal As Label
    Private _lblResumen As Label
    Private _btnMarcarTodos As Button
    Private _btnDesmarcarTodos As Button
    Private _btnAceptar As Button
    Private _btnCancelar As Button

    Public ReadOnly Property Items As List(Of ElementoMecItem)
        Get
            Return _items
        End Get
    End Property

    Public Sub New(items As List(Of ElementoMecItem))
        _items = items
        InitializeComponent()
        CargarDatos()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Sobremedida de Largo (+3mm) - Revisión General de Componentes"
        Me.Size = New Size(980, 600)
        Me.MinimumSize = New Size(800, 480)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Font = New Font("Segoe UI", 9.0F)

        ' =========================================================
        ' PANEL SUPERIOR: Instrucciones y Botones de Selección Rápida
        ' =========================================================
        Dim pnlTop As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 70,
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
            .Text = "Seleccione las piezas a las que desea sumar +3 mm de sobremedida en la columna 'SUMA'. Use el desplegable para ver la miniatura 3D.",
            .ForeColor = System.Drawing.Color.FromArgb(100, 116, 139),
            .Location = New System.Drawing.Point(12, 28),
            .AutoSize = True
        }

        _btnMarcarTodos = New Button With {
            .Text = "☑ Marcar Todos (+3mm)",
            .Location = New System.Drawing.Point(12, 45),
            .Size = New Size(160, 24),
            .BackColor = System.Drawing.Color.White
        }
        AddHandler _btnMarcarTodos.Click, AddressOf OnMarcarTodos

        _btnDesmarcarTodos = New Button With {
            .Text = "☐ Desmarcar Todos",
            .Location = New System.Drawing.Point(178, 45),
            .Size = New Size(140, 24),
            .BackColor = System.Drawing.Color.White
        }
        AddHandler _btnDesmarcarTodos.Click, AddressOf OnDesmarcarTodos

        Dim lblCmb As New Label With {
            .Text = "Vista previa desplegable:",
            .Location = New System.Drawing.Point(450, 48),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        }

        _cmbPiezas = New ComboBox With {
            .Location = New System.Drawing.Point(600, 45),
            .Size = New Size(350, 24),
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
        ' PANEL DERECHO: Vista Previa y Ficha Técnica
        ' =========================================================
        Dim pnlRight As New Panel With {
            .Dock = DockStyle.Right,
            .Width = 240,
            .Padding = New Padding(10),
            .BackColor = System.Drawing.Color.FromArgb(241, 245, 249)
        }

        Dim grpPreview As New GroupBox With {
            .Text = "Vista Previa de la Pieza",
            .Dock = DockStyle.Fill,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(30, 41, 59)
        }

        _picPreview = New PictureBox With {
            .Location = New System.Drawing.Point(15, 22),
            .Size = New Size(190, 160),
            .SizeMode = PictureBoxSizeMode.Zoom,
            .BorderStyle = BorderStyle.FixedSingle,
            .BackColor = System.Drawing.Color.White
        }

        _lblPartNumber = New Label With {
            .Location = New System.Drawing.Point(15, 195),
            .Size = New Size(190, 36),
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(2, 132, 199),
            .Text = "Part Number: -"
        }

        _lblStockNumber = New Label With {
            .Location = New System.Drawing.Point(15, 235),
            .Size = New Size(190, 20),
            .Font = New Font("Segoe UI", 8.0F),
            .Text = "Stock Number: -"
        }

        _lblMaterial = New Label With {
            .Location = New System.Drawing.Point(15, 260),
            .Size = New Size(190, 35),
            .Font = New Font("Segoe UI", 8.0F),
            .Text = "Material: -"
        }

        _lblLargoFinal = New Label With {
            .Location = New System.Drawing.Point(15, 300),
            .Size = New Size(190, 40),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(16, 185, 129),
            .Text = "Largo Final: -"
        }

        grpPreview.Controls.Add(_picPreview)
        grpPreview.Controls.Add(_lblPartNumber)
        grpPreview.Controls.Add(_lblStockNumber)
        grpPreview.Controls.Add(_lblMaterial)
        grpPreview.Controls.Add(_lblLargoFinal)
        pnlRight.Controls.Add(grpPreview)

        ' =========================================================
        ' PANEL INFERIOR: Resumen y Botones Aceptar / Cancelar
        ' =========================================================
        Dim pnlBottom As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 48,
            .Padding = New Padding(12, 10, 12, 10),
            .BackColor = System.Drawing.Color.FromArgb(248, 250, 252)
        }

        _lblResumen = New Label With {
            .Location = New System.Drawing.Point(12, 15),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(51, 65, 85),
            .Text = "Total componentes: 0"
        }

        _btnAceptar = New Button With {
            .Text = "Aceptar y Aplicar",
            .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom,
            .Location = New System.Drawing.Point(730, 10),
            .Size = New Size(125, 28),
            .BackColor = System.Drawing.Color.FromArgb(2, 132, 199),
            .ForeColor = System.Drawing.Color.White,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler _btnAceptar.Click, AddressOf OnAceptar

        _btnCancelar = New Button With {
            .Text = "Cancelar",
            .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom,
            .Location = New System.Drawing.Point(865, 10),
            .Size = New Size(85, 28),
            .BackColor = System.Drawing.Color.White
        }
        AddHandler _btnCancelar.Click, AddressOf OnCancelar

        pnlBottom.Controls.Add(_lblResumen)
        pnlBottom.Controls.Add(_btnAceptar)
        pnlBottom.Controls.Add(_btnCancelar)

        ' =========================================================
        ' CENTRO: DataGridView con todas las columnas
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
            .Font = New Font("Segoe UI", 8.5F)
        }

        ConfigurarColumnasGrid()
        AddHandler _grid.CellValueChanged, AddressOf OnGridCellValueChanged
        AddHandler _grid.CurrentCellDirtyStateChanged, AddressOf OnGridCurrentCellDirtyStateChanged
        AddHandler _grid.SelectionChanged, AddressOf OnGridSelectionChanged

        ' Agregar paneles al Form
        Me.Controls.Add(_grid)
        Me.Controls.Add(pnlRight)
        Me.Controls.Add(pnlTop)
        Me.Controls.Add(pnlBottom)

        Me.AcceptButton = _btnAceptar
        Me.CancelButton = _btnCancelar
    End Sub

    Private Sub ConfigurarColumnasGrid()
        ' Columna 1: SUMA (+3mm) CheckBox
        Dim colSuma As New DataGridViewCheckBoxColumn()
        colSuma.HeaderText = "SUMA (+3mm)"
        colSuma.Name = "colSuma"
        colSuma.Width = 95
        colSuma.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        colSuma.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        ' Columna 2: Part Number
        Dim colPart As New DataGridViewTextBoxColumn With {
            .HeaderText = "Part Number",
            .Name = "colPart",
            .ReadOnly = True,
            .Width = 140
        }

        ' Columna 3: Stock Number
        Dim colStock As New DataGridViewTextBoxColumn With {
            .HeaderText = "Stock Number",
            .Name = "colStock",
            .ReadOnly = True,
            .Width = 110
        }

        ' Columna 4: Material Plano
        Dim colMat As New DataGridViewTextBoxColumn With {
            .HeaderText = "Material Plano",
            .Name = "colMat",
            .ReadOnly = True,
            .Width = 125
        }

        ' Columna 5: Largo Original
        Dim colLargoOrig As New DataGridViewTextBoxColumn With {
            .HeaderText = "Largo Orig.",
            .Name = "colLargoOrig",
            .ReadOnly = True,
            .Width = 90,
            .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight}
        }

        ' Columna 6: Largo Final (+3mm)
        Dim colLargoFinal As New DataGridViewTextBoxColumn With {
            .HeaderText = "Largo Final",
            .Name = "colLargoFinal",
            .ReadOnly = True,
            .Width = 95,
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
            .Width = 55,
            .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleCenter}
        }

        _grid.Columns.AddRange(New DataGridViewColumn() {
            colSuma, colPart, colStock, colMat, colLargoOrig, colLargoFinal, colAncho, colTipo
        })
    End Sub

    Private Sub CargarDatos()
        _grid.Rows.Clear()
        _cmbPiezas.Items.Clear()

        For Each item In _items
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
            _cmbPiezas.Items.Add(item)
        Next

        If _cmbPiezas.Items.Count > 0 Then
            _cmbPiezas.SelectedIndex = 0
        End If

        ActualizarResumen()
    End Sub

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

        ' Si es la fila seleccionada, actualizar ficha derecha
        If _grid.CurrentRow IsNot Nothing AndAlso _grid.CurrentRow.Index = rowIdx Then
            _lblLargoFinal.Text = "Largo Final: " & item.LargoFinalStr & If(item.Sumar3mm, " (+3mm)", " (Nominal)")
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

    Private Sub MostrarDetalleItem(item As ElementoMecItem)
        If item Is Nothing Then
            _picPreview.Image = Nothing
            _lblPartNumber.Text = "Part Number: -"
            _lblStockNumber.Text = "Stock Number: -"
            _lblMaterial.Text = "Material: -"
            _lblLargoFinal.Text = "Largo Final: -"
            Exit Sub
        End If

        _picPreview.Image = item.Thumbnail
        _lblPartNumber.Text = "Part Number: " & item.PartNumber
        _lblStockNumber.Text = "Stock Number: " & item.StockNumber
        _lblMaterial.Text = "Material: " & item.MaterialPlano
        _lblLargoFinal.Text = "Largo Final: " & item.LargoFinalStr & If(item.Sumar3mm, " (+3mm)", " (Nominal)")
    End Sub

    ' =========================================================
    ' EVENTOS DE INTERFAZ
    ' =========================================================

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
                item.Sumar3mm = CBool(_grid.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
                ActualizarFilaLargo(e.RowIndex)
                ActualizarResumen()
            End If
        End If
    End Sub

    Private Sub OnGridSelectionChanged(sender As Object, e As EventArgs)
        If _grid.CurrentRow IsNot Nothing AndAlso _grid.CurrentRow.Tag IsNot Nothing Then
            Dim item = DirectCast(_grid.CurrentRow.Tag, ElementoMecItem)
            MostrarDetalleItem(item)

            ' Sincronizar ComboBox desplegable
            If _cmbPiezas.SelectedItem IsNot item Then
                _cmbPiezas.SelectedItem = item
            End If
        End If
    End Sub

    Private Sub OnCmbSelectedIndexChanged(sender As Object, e As EventArgs)
        Dim item = DirectCast(_cmbPiezas.SelectedItem, ElementoMecItem)
        If item IsNot Nothing Then
            MostrarDetalleItem(item)

            ' Sincronizar selección en el DataGridView
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

    Private Sub OnAceptar(sender As Object, e As EventArgs)
        DialogResult = DialogResult.OK
        Close()
    End Sub

    Private Sub OnCancelar(sender As Object, e As EventArgs)
        DialogResult = DialogResult.Cancel
        Close()
    End Sub

End Class
