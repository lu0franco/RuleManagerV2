Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Windows.Forms
Imports Inventor
Imports IO = System.IO

''' <summary>
''' Diálogo interactivo para revisar y ajustar manualmente o automáticamente las cantidades en una Lista de Piezas (PartsList).
''' 
''' Permite ver:
''' - La cantidad actual en la tabla (columna CANT / CANTIDAD_USADA)
''' - La cantidad de ocurrencias reales en el subensamblaje del plano (ej: 4 ocurrencias de "PLATINA DE ANCLAJE" en CHASIS INFERIOR)
''' - Un multiplicador de fabricación (ej: 2 chasis -> 4 * 2 = 8 platinas)
''' - Edición manual libre directamente sobre la celda
''' - Reemplazo del valor estático de la celda (cell.Static = True, cell.Value = nuevaCantidad)
'''   sin alterar las iProperties del modelo 3D para que el plano individual de fabricación siga marcando el total (ej: 16).
''' </summary>
Public Class AjustarCantidadesDialog
    Inherits Form

    Private ReadOnly _invApp As Inventor.Application
    Private ReadOnly _partsList As PartsList
    Private ReadOnly _sheet As Sheet
    Private ReadOnly _asmDoc As AssemblyDocument
    Private ReadOnly _colCantIndex As Integer

    Private _grid As DataGridView
    Private _lblTituloInfo As Label
    Private _numMultiplicador As NumericUpDown
    Private _btnMultiplicarOcurrencias As Button
    Private _btnCopiarOcurrencias1a1 As Button
    Private _btnAplicar As Button
    Private _btnRestaurar As Button
    Private _btnCancelar As Button
    Private _lblResumen As Label
    Private _cmbColumna As ComboBox

    Public Property FilasModificadas As Integer = 0

    Public Sub New(invApp As Inventor.Application, pl As PartsList, sheet As Sheet, asmDoc As AssemblyDocument, colCantIndex As Integer)
        _invApp = invApp
        _partsList = pl
        _sheet = sheet
        _asmDoc = asmDoc
        _colCantIndex = colCantIndex

        InitializeComponent()
        CargarDatos()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "RuleManager V2 - Ajuste de Cantidades en Lista de Piezas (PartsList)"
        Me.Size = New Size(1020, 640)
        Me.MinimumSize = New Size(850, 500)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.MinimizeBox = False
        Me.MaximizeBox = True
        Me.Font = New Font("Segoe UI", 9.0F)
        Me.BackColor = System.Drawing.Color.FromArgb(248, 249, 250)

        ' Header Panel
        Dim pnlHeader As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 72,
            .BackColor = System.Drawing.Color.FromArgb(15, 23, 42),
            .Padding = New Padding(16, 10, 16, 10)
        }

        Dim lblTitle As New Label With {
            .Text = "Ajuste de Cantidades por Subensamblaje (Columna CANT / CANTIDAD_USADA)",
            .ForeColor = System.Drawing.Color.White,
            .Font = New Font("Segoe UI", 11.0F, FontStyle.Bold),
            .AutoSize = True,
            .Location = New System.Drawing.Point(16, 10)
        }

        Dim asmName As String = If(_asmDoc IsNot Nothing, IO.Path.GetFileName(_asmDoc.FullFileName), "Subensamblaje del Plano")
        Dim colTitle As String = "CANT"
        Try
            colTitle = _partsList.PartsListColumns.Item(_colCantIndex).Title
        Catch
        End Try

        _lblTituloInfo = New Label With {
            .Text = $"Hoja: {_sheet.Name}   |   Subensamblaje 3D: {asmName}   |   Columna editada: {colTitle} (Col {_colCantIndex})",
            .ForeColor = System.Drawing.Color.FromArgb(148, 163, 184),
            .Font = New Font("Segoe UI", 8.5F),
            .AutoSize = True,
            .Location = New System.Drawing.Point(16, 38)
        }

        pnlHeader.Controls.AddRange({lblTitle, _lblTituloInfo})

        ' Top Toolbar Panel
        Dim pnlToolbar As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 52,
            .BackColor = System.Drawing.Color.White,
            .Padding = New Padding(16, 10, 16, 10)
        }

        Dim lblMult As New Label With {
            .Text = "Multiplicador de fabricación:",
            .AutoSize = True,
            .Location = New System.Drawing.Point(16, 16),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(30, 41, 59)
        }

        _numMultiplicador = New NumericUpDown With {
            .Location = New System.Drawing.Point(190, 13),
            .Width = 55,
            .Minimum = 1,
            .Maximum = 100,
            .Value = 1,
            .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        }
        AddHandler _numMultiplicador.ValueChanged, AddressOf OnMultiplicadorChanged

        _btnMultiplicarOcurrencias = New Button With {
            .Text = "⚡ Calcular (Ocurrencias × Mult.)",
            .Location = New System.Drawing.Point(260, 11),
            .Size = New Size(210, 30),
            .BackColor = System.Drawing.Color.FromArgb(224, 231, 255),
            .ForeColor = System.Drawing.Color.FromArgb(49, 46, 129),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .Cursor = Cursors.Hand
        }
        _btnMultiplicarOcurrencias.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(165, 180, 252)
        AddHandler _btnMultiplicarOcurrencias.Click, AddressOf OnMultiplicarOcurrenciasClicked

        _btnCopiarOcurrencias1a1 = New Button With {
            .Text = "Copiar Ocurrencias 1:1",
            .Location = New System.Drawing.Point(480, 11),
            .Size = New Size(160, 30),
            .BackColor = System.Drawing.Color.FromArgb(241, 245, 249),
            .ForeColor = System.Drawing.Color.FromArgb(51, 65, 85),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 8.5F),
            .Cursor = Cursors.Hand
        }
        _btnCopiarOcurrencias1a1.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(203, 213, 225)
        AddHandler _btnCopiarOcurrencias1a1.Click, AddressOf OnCopiarOcurrencias1a1Clicked

        _lblResumen = New Label With {
            .AutoSize = True,
            .Location = New System.Drawing.Point(660, 16),
            .ForeColor = System.Drawing.Color.FromArgb(71, 85, 105),
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Italic)
        }

        pnlToolbar.Controls.AddRange({lblMult, _numMultiplicador, _btnMultiplicarOcurrencias, _btnCopiarOcurrencias1a1, _lblResumen})

        ' Grid Panel
        Dim pnlGrid As New Panel With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(16, 8, 16, 8)
        }

        _grid = New DataGridView With {
            .Dock = DockStyle.Fill,
            .BackgroundColor = System.Drawing.Color.White,
            .BorderStyle = BorderStyle.FixedSingle,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .SelectionMode = DataGridViewSelectionMode.CellSelect,
            .RowHeadersVisible = False,
            .Font = New Font("Segoe UI", 9.0F),
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .EditMode = DataGridViewEditMode.EditOnEnter
        }
        AddHandler _grid.CellValueChanged, AddressOf OnGridCellValueChanged

        pnlGrid.Controls.Add(_grid)

        ' Bottom Panel
        Dim pnlBottom As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 56,
            .BackColor = System.Drawing.Color.FromArgb(241, 245, 249),
            .Padding = New Padding(16, 12, 16, 12)
        }

        _btnRestaurar = New Button With {
            .Text = "Restaurar Dinámicos (Quitar Override Estático)",
            .Dock = DockStyle.Left,
            .Width = 320,
            .BackColor = System.Drawing.Color.White,
            .FlatStyle = FlatStyle.Flat,
            .ForeColor = System.Drawing.Color.FromArgb(71, 85, 105),
            .Cursor = Cursors.Hand
        }
        _btnRestaurar.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(203, 213, 225)
        AddHandler _btnRestaurar.Click, AddressOf OnRestaurarClicked

        _btnCancelar = New Button With {
            .Text = "Cancelar",
            .Dock = DockStyle.Right,
            .Width = 100,
            .BackColor = System.Drawing.Color.White,
            .FlatStyle = FlatStyle.Flat,
            .ForeColor = System.Drawing.Color.FromArgb(51, 65, 85),
            .DialogResult = DialogResult.Cancel,
            .Cursor = Cursors.Hand
        }
        _btnCancelar.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(203, 213, 225)

        _btnAplicar = New Button With {
            .Text = "Aplicar a Lista de Piezas",
            .Dock = DockStyle.Right,
            .Width = 200,
            .BackColor = System.Drawing.Color.FromArgb(2, 132, 199),
            .ForeColor = System.Drawing.Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .Cursor = Cursors.Hand
        }
        _btnAplicar.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(3, 105, 161)
        AddHandler _btnAplicar.Click, AddressOf OnAplicarClicked

        pnlBottom.Controls.AddRange({_btnRestaurar, _btnAplicar, _btnCancelar})

        Me.Controls.AddRange({pnlGrid, pnlToolbar, pnlBottom, pnlHeader})
    End Sub

    Private Sub CargarDatos()
        _grid.Columns.Clear()

        ' Configuración de columnas utilizando FillWeight para repartir el espacio
        _grid.Columns.Add(New DataGridViewTextBoxColumn With {
        .Name = "colPos",
        .HeaderText = "POS",
        .FillWeight = 8,
        .ReadOnly = True
    })

        _grid.Columns.Add(New DataGridViewTextBoxColumn With {
        .Name = "colDesc",
        .HeaderText = "DESCRIPCIÓN",
        .FillWeight = 25,
        .ReadOnly = True
    })

        _grid.Columns.Add(New DataGridViewTextBoxColumn With {
        .Name = "colDetalle",
        .HeaderText = "DETALLE / MATERIAL / N° PIEZA",
        .FillWeight = 25,
        .ReadOnly = True
    })

        _grid.Columns.Add(New DataGridViewTextBoxColumn With {
        .Name = "colCantActual",
        .HeaderText = "CANT ACTUAL (TABLA)",
        .FillWeight = 12,
        .ReadOnly = True
    })

        _grid.Columns.Add(New DataGridViewTextBoxColumn With {
        .Name = "colOcurrencias3D",
        .HeaderText = "OCURRENCIAS 3D",
        .FillWeight = 12,
        .ReadOnly = True
    })

        _grid.Columns.Add(New DataGridViewTextBoxColumn With {
        .Name = "colNuevaCant",
        .HeaderText = "NUEVA CANTIDAD (EDITABLE)",
        .FillWeight = 18,
        .ReadOnly = False
    })

        _grid.Columns.Add(New DataGridViewTextBoxColumn With {
        .Name = "colEstado",
        .HeaderText = "ESTADO",
        .FillWeight = 20,
        .ReadOnly = True
    })

        ' Estilizar columna editable
        _grid.Columns("colNuevaCant").DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(254, 249, 195)
        _grid.Columns("colNuevaCant").DefaultCellStyle.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        _grid.Columns("colNuevaCant").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        _grid.Columns("colCantActual").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        _grid.Columns("colOcurrencias3D").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        _grid.Columns("colPos").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        ' Resto de tu código (ObtenerConteosEnsamblaje, rellenar filas, etc.) ...

        ' 1. Obtener conteo de ocurrencias reales del ensamblaje del plano
        Dim conteos3D As Dictionary(Of String, Integer) = ObtenerConteosEnsamblaje(_asmDoc)

        ' 2. Identificar columnas de descripción y detalle/part number en la PartsList
        Dim colDescIdx As Integer = -1
        Dim colDetalleIdx As Integer = -1

        For cIdx As Integer = 1 To _partsList.PartsListColumns.Count
            Dim titleUpper As String = _partsList.PartsListColumns.Item(cIdx).Title.Trim().ToUpper()
            If titleUpper.Contains("DESC") Then colDescIdx = cIdx
            If titleUpper.Contains("DETALLE") OrElse
               titleUpper.Contains("MATERIAL") OrElse
               titleUpper.Contains("PART NUMBER") OrElse
               titleUpper.Contains("PIEZA") OrElse
               titleUpper.Contains("Nº") OrElse
               titleUpper.Contains("ALMACENAMIENTO") Then
                If colDetalleIdx = -1 Then colDetalleIdx = cIdx
            End If
        Next

        Dim filasDiferentes As Integer = 0
        Dim mult As Integer = CInt(_numMultiplicador.Value)

        ' 3. Rellenar DataGridView
        For rIdx As Integer = 1 To _partsList.PartsListRows.Count
            Dim row As PartsListRow = _partsList.PartsListRows.Item(rIdx)
            If row.Custom Then Continue For

            Dim posVal As String = rIdx.ToString()
            Try
                If row.Item(1).Value IsNot Nothing Then
                    posVal = row.Item(1).Value.ToString().Trim()
                End If
            Catch
            End Try

            Dim descVal As String = ""
            If colDescIdx > 0 AndAlso row.Item(colDescIdx).Value IsNot Nothing Then
                descVal = row.Item(colDescIdx).Value.ToString().Trim()
            End If

            Dim detalleVal As String = ""
            If colDetalleIdx > 0 AndAlso row.Item(colDetalleIdx).Value IsNot Nothing Then
                detalleVal = row.Item(colDetalleIdx).Value.ToString().Trim()
            End If

            Dim cantActual As String = ""
            If row.Item(_colCantIndex).Value IsNot Nothing Then
                cantActual = row.Item(_colCantIndex).Value.ToString().Trim()
            End If

            ' Buscar conteo en el diccionario 3D usando Detalle, Descripción y Part Number
            Dim ocurrs3D As Integer = BuscarConteo(conteos3D, detalleVal, descVal)

            Dim nuevoVal As String = cantActual
            Dim estado As String = "Sin cambio"

            ' Comprobar si difiere del conteo multiplicado
            Dim nActual As Double = 0
            Dim esNum As Boolean = Double.TryParse(cantActual, nActual)

            If ocurrs3D > 0 Then
                Dim targetCalculado As Integer = ocurrs3D * mult
                If esNum AndAlso Math.Abs(nActual - targetCalculado) > 0.001 Then
                    nuevoVal = targetCalculado.ToString()
                    estado = $"⚠️ Difiere de 3D ({nActual} vs {targetCalculado})"
                    filasDiferentes += 1
                ElseIf row.Item(_colCantIndex).Static Then
                    estado = "Fijado Estático"
                End If
            ElseIf row.Item(_colCantIndex).Static Then
                estado = "Fijado Estático"
            End If

            Dim rowIndex As Integer = _grid.Rows.Add(
                posVal,
                descVal,
                detalleVal,
                cantActual,
                If(ocurrs3D > 0, ocurrs3D.ToString(), "-"),
                nuevoVal,
                estado
            )

            ' Resaltar estado
            If estado.StartsWith("⚠️") Then
                _grid.Rows(rowIndex).Cells("colCantActual").Style.ForeColor = System.Drawing.Color.Crimson
                _grid.Rows(rowIndex).Cells("colCantActual").Style.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
                _grid.Rows(rowIndex).Cells("colEstado").Style.ForeColor = System.Drawing.Color.DarkOrange
                _grid.Rows(rowIndex).Cells("colEstado").Style.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
            ElseIf estado = "Fijado Estático" Then
                _grid.Rows(rowIndex).Cells("colEstado").Style.ForeColor = System.Drawing.Color.FromArgb(2, 132, 199)
            End If
        Next

        _lblResumen.Text = $"{_grid.Rows.Count} componentes en lista  |  {filasDiferentes} con discrepancia de cantidad"
    End Sub

    Private Function ObtenerConteosEnsamblaje(asmDoc As AssemblyDocument) As Dictionary(Of String, Integer)
        Dim dict As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        If asmDoc Is Nothing Then Return dict

        Try
            ContarOcurrenciasRecursivas(asmDoc.ComponentDefinition.Occurrences, dict)
        Catch ex As Exception
        End Try

        Return dict
    End Function

    Private Sub ContarOcurrenciasRecursivas(occs As ComponentOccurrences, dict As Dictionary(Of String, Integer))
        For Each occ As ComponentOccurrence In occs
            If occ.Suppressed Then Continue For
            If occ.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then Continue For

            Try
                Dim doc As Document = occ.Definition.Document
                If doc IsNot Nothing Then
                    RegistrarDocEnConteo(dict, doc)
                End If
            Catch
            End Try

            ' Si es subensamblaje, recorrer sus componentes internos
            If occ.SubOccurrences IsNot Nothing AndAlso occ.SubOccurrences.Count > 0 Then
                ContarOcurrenciasRecursivas(occ.SubOccurrences, dict)
            End If
        Next
    End Sub

    Private Sub RegistrarDocEnConteo(dict As Dictionary(Of String, Integer), doc As Document)
        Dim pn As String = ""
        Dim stockNum As String = ""
        Dim desc As String = ""

        Try
            Dim dtProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
            Try : pn = dtProps.Item("Part Number").Value?.ToString().Trim() : Catch : End Try
            Try : stockNum = dtProps.Item("Stock Number").Value?.ToString().Trim() : Catch : End Try
            Try : desc = dtProps.Item("Description").Value?.ToString().Trim() : Catch : End Try
        Catch
        End Try

        Dim docDisplayName As String = doc.DisplayName.Trim()
        Dim docFileName As String = IO.Path.GetFileNameWithoutExtension(doc.FullFileName).Trim()

        ' Normalizar simetrías quitando sufijo _MIR
        Dim pnBase As String = pn
        If pnBase.ToUpper().EndsWith("_MIR") Then
            pnBase = pnBase.Substring(0, pnBase.Length - 4)
        End If

        If Not String.IsNullOrEmpty(pn) Then AgregarConteo(dict, pn)
        If pnBase <> pn AndAlso Not String.IsNullOrEmpty(pnBase) Then AgregarConteo(dict, pnBase)
        If Not String.IsNullOrEmpty(stockNum) Then AgregarConteo(dict, stockNum)
        If Not String.IsNullOrEmpty(desc) AndAlso Not desc.Equals("SIMETRIA", StringComparison.OrdinalIgnoreCase) Then
            AgregarConteo(dict, desc)
        End If
        If Not String.IsNullOrEmpty(docDisplayName) Then AgregarConteo(dict, docDisplayName)
        If Not String.IsNullOrEmpty(docFileName) Then AgregarConteo(dict, docFileName)
    End Sub

    Private Sub AgregarConteo(dict As Dictionary(Of String, Integer), clave As String)
        If String.IsNullOrWhiteSpace(clave) Then Exit Sub
        Dim k As String = clave.Trim()
        If dict.ContainsKey(k) Then
            dict(k) += 1
        Else
            dict(k) = 1
        End If
    End Sub

    Private Function BuscarConteo(dict As Dictionary(Of String, Integer), ParamArray claves As String()) As Integer
        For Each c In claves
            If String.IsNullOrWhiteSpace(c) Then Continue For
            Dim clean As String = c.Trim()

            ' 1. Coincidencia exacta
            If dict.ContainsKey(clean) Then Return dict(clean)

            ' 2. Sin extensión
            If clean.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) OrElse clean.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then
                Dim noExt As String = IO.Path.GetFileNameWithoutExtension(clean)
                If dict.ContainsKey(noExt) Then Return dict(noExt)
            End If

            ' 3. Sin sufijo _MIR
            If clean.ToUpper().EndsWith("_MIR") Then
                Dim noMir As String = clean.Substring(0, clean.Length - 4)
                If dict.ContainsKey(noMir) Then Return dict(noMir)
            End If
        Next

        ' 4. Búsqueda por contención o parte de texto
        For Each c In claves
            If String.IsNullOrWhiteSpace(c) Then Continue For
            Dim cleanUpper As String = c.Trim().ToUpper()
            If cleanUpper.Length < 3 Then Continue For

            For Each kv In dict
                Dim kUpper As String = kv.Key.ToUpper()
                If kUpper = cleanUpper OrElse kUpper.Contains(cleanUpper) OrElse cleanUpper.Contains(kUpper) Then
                    Return kv.Value
                End If
            Next
        Next

        Return 0
    End Function

    Private Sub OnMultiplicadorChanged(sender As Object, e As EventArgs)
        RecalcularCantidades()
    End Sub

    Private Sub OnMultiplicarOcurrenciasClicked(sender As Object, e As EventArgs)
        RecalcularCantidades()
        MessageBox.Show($"Se recalculó 'NUEVA CANTIDAD' aplicando el multiplicador ({_numMultiplicador.Value}x) sobre las ocurrencias 3D." & vbCrLf &
                        "Puedes ajustar manualmente cualquier valor antes de presionar 'Aplicar'.",
                        "Recálculo Completado", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub OnCopiarOcurrencias1a1Clicked(sender As Object, e As EventArgs)
        For Each row As DataGridViewRow In _grid.Rows
            Dim ocurrStr As String = row.Cells("colOcurrencias3D").Value?.ToString()
            Dim ocurrVal As Integer = 0
            If Integer.TryParse(ocurrStr, ocurrVal) AndAlso ocurrVal > 0 Then
                row.Cells("colNuevaCant").Value = ocurrVal.ToString()
                row.Cells("colEstado").Value = "Listo (1:1)"
                row.Cells("colEstado").Style.ForeColor = System.Drawing.Color.Green
            End If
        Next
    End Sub

    Private Sub RecalcularCantidades()
        Dim mult As Integer = CInt(_numMultiplicador.Value)
        For Each row As DataGridViewRow In _grid.Rows
            Dim ocurrStr As String = row.Cells("colOcurrencias3D").Value?.ToString()
            Dim ocurrVal As Integer = 0
            If Integer.TryParse(ocurrStr, ocurrVal) AndAlso ocurrVal > 0 Then
                Dim targetVal As Integer = ocurrVal * mult
                row.Cells("colNuevaCant").Value = targetVal.ToString()
                row.Cells("colEstado").Value = $"Listo ({mult}x)"
                row.Cells("colEstado").Style.ForeColor = System.Drawing.Color.Green
                row.Cells("colEstado").Style.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
            End If
        Next
    End Sub

    Private Sub OnGridCellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Exit Sub
        If _grid.Columns(e.ColumnIndex).Name = "colNuevaCant" Then
            Dim row As DataGridViewRow = _grid.Rows(e.RowIndex)
            Dim cantActual As String = row.Cells("colCantActual").Value?.ToString().Trim()
            Dim nuevaCant As String = row.Cells("colNuevaCant").Value?.ToString().Trim()

            If Not String.IsNullOrEmpty(nuevaCant) AndAlso Not nuevaCant.Equals(cantActual, StringComparison.OrdinalIgnoreCase) Then
                row.Cells("colEstado").Value = "Modificado manual"
                row.Cells("colEstado").Style.ForeColor = System.Drawing.Color.Green
                row.Cells("colEstado").Style.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
            End If
        End If
    End Sub

    Private Sub OnAplicarClicked(sender As Object, e As EventArgs)
        Dim modificadas As Integer = 0

        For rIdx As Integer = 0 To _grid.Rows.Count - 1
            Dim rowGrid As DataGridViewRow = _grid.Rows(rIdx)
            Dim cantActual As String = rowGrid.Cells("colCantActual").Value?.ToString().Trim()
            Dim nuevaCant As String = rowGrid.Cells("colNuevaCant").Value?.ToString().Trim()

            If Not String.IsNullOrEmpty(nuevaCant) AndAlso Not nuevaCant.Equals(cantActual, StringComparison.OrdinalIgnoreCase) Then
                ' Aplicar override estático en la celda de la lista de piezas
                Dim plRow As PartsListRow = _partsList.PartsListRows.Item(rIdx + 1)
                Dim cell As PartsListCell = plRow.Item(_colCantIndex)

                cell.Static = True
                cell.Value = nuevaCant
                modificadas += 1
            End If
        Next

        Me.FilasModificadas = modificadas

        If modificadas > 0 Then
            Try
                Dim drawDoc As DrawingDocument = DirectCast(_sheet.Parent, DrawingDocument)
                drawDoc.Update2(True)
                _invApp.ActiveView.Update()
            Catch
            End Try

            MessageBox.Show($"Se aplicó el reemplazo estático a {modificadas} fila(s) en la columna CANT:" & vbCrLf & vbCrLf &
                            "✓ La celda ahora muestra el valor requerido para este subensamblaje." & vbCrLf &
                            "✓ La iProperty 3D 'CANTIDAD_USADA' del archivo de pieza permanece intacta para su fabricación.",
                            "Ajuste Completado", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("No se realizaron modificaciones porque los valores no han cambiado.",
                            "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub OnRestaurarClicked(sender As Object, e As EventArgs)
        Dim resp As DialogResult = MessageBox.Show(
            "¿Desea quitar el valor estático (override) de todas las celdas de CANT en esta lista de piezas para que vuelvan a leer la propiedad dinámica del modelo 3D?",
            "Restaurar Valores Dinámicos",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question)

        If resp <> DialogResult.Yes Then Exit Sub

        Dim restauradas As Integer = 0
        For rIdx As Integer = 1 To _partsList.PartsListRows.Count
            Dim plRow As PartsListRow = _partsList.PartsListRows.Item(rIdx)
            Dim cell As PartsListCell = plRow.Item(_colCantIndex)
            If cell.Static Then
                cell.Static = False
                restauradas += 1
            End If
        Next

        Try
            Dim drawDoc As DrawingDocument = DirectCast(_sheet.Parent, DrawingDocument)
            drawDoc.Update2(True)
            _invApp.ActiveView.Update()
        Catch
        End Try

        MessageBox.Show($"Se restauraron {restauradas} celdas al valor original dinámico del modelo.",
                        "Restauración Completada", MessageBoxButtons.OK, MessageBoxIcon.Information)

        CargarDatos()
    End Sub

End Class
