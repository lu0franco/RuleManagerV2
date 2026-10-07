Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor
Imports IO = System.IO

Public Class ReviewDialog
    Inherits Form

    Private _allItems As List(Of FileMoveItem)
    Private _rootPath As String
    Private _invApp As Inventor.Application
    Private _showExcluded As Boolean = False
    Private _sortColumn As Integer = -1
    Private _sortAscending As Boolean = True

    Private WithEvents _grid As DataGridView
    Private _btnToggleExcluded As Button
    Private _lblInfo As Label
    Private _progress As System.Windows.Forms.ProgressBar

    ' Referencias a los paneles para calcular layout
    Private _topPanel As Panel
    Private _bottomPanel As Panel
    Private _gridContainer As Panel

    Private ReadOnly _carpetasProyecto As New List(Of String) From {
        "01 - DIBUJOS", "02 - PLANOS", "03 - CORTES",
        "04 - COMERCIALES", "05 - PIEZAS COMUNES",
        "06 - MULTIMEDIA", "07 - AUTOCAD - SAT",
        "08 - MAQUINAS DE REFERENCIAS"
    }

    Public Sub New(items As List(Of FileMoveItem), rootPath As String, invApp As Inventor.Application)
        _allItems = items
        _rootPath = rootPath
        _invApp = invApp
        InitializeDialog()
    End Sub

    Private Sub InitializeDialog()
        Me.Text = "Revisar organización de archivos"
        Me.Size = New System.Drawing.Size(1150, 720)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.MinimizeBox = False
        Me.MaximizeBox = True

        ' =========================================================
        ' PANEL SUPERIOR (Dock.Top)
        ' =========================================================
        _topPanel = New Panel()
        _topPanel.Dock = DockStyle.Top
        _topPanel.Height = 42
        _topPanel.Padding = New Padding(10, 8, 10, 4)

        _btnToggleExcluded = New Button()
        _btnToggleExcluded.FlatStyle = FlatStyle.Flat
        _btnToggleExcluded.Height = 28
        _btnToggleExcluded.AutoSize = True
        _btnToggleExcluded.TextAlign = ContentAlignment.MiddleLeft
        _btnToggleExcluded.Font = New System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold)
        AddHandler _btnToggleExcluded.Click, AddressOf ToggleExcluded_Click

        _lblInfo = New Label()
        _lblInfo.AutoSize = True
        _lblInfo.Text = "Doble clic en 'Destino propuesto' para cambiar la carpeta. Presioná Aceptar para ejecutar."
        _lblInfo.ForeColor = System.Drawing.Color.Gray
        _lblInfo.Font = New System.Drawing.Font("Segoe UI", 9)

        _topPanel.Controls.Add(_btnToggleExcluded)
        _topPanel.Controls.Add(_lblInfo)
        _lblInfo.Location = New System.Drawing.Point(320, 12)

        ' =========================================================
        ' PANEL INFERIOR (Dock.Bottom) - Botones
        ' =========================================================
        _bottomPanel = New Panel()
        _bottomPanel.Dock = DockStyle.Bottom
        _bottomPanel.Height = 55
        _bottomPanel.Padding = New Padding(10)

        Dim btnOk As New Button()
        btnOk.Text = "Aceptar y organizar"
        btnOk.Size = New System.Drawing.Size(140, 30)
        btnOk.Location = New System.Drawing.Point(_bottomPanel.Width - 270, 12)
        btnOk.Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
        AddHandler btnOk.Click, AddressOf BtnOk_Click

        Dim btnCancel As New Button()
        btnCancel.Text = "Cancelar"
        btnCancel.Size = New System.Drawing.Size(90, 30)
        btnCancel.Location = New System.Drawing.Point(_bottomPanel.Width - 120, 12)
        btnCancel.Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
        AddHandler btnCancel.Click, Sub(s, e) Me.DialogResult = DialogResult.Cancel

        _bottomPanel.Controls.Add(btnOk)
        _bottomPanel.Controls.Add(btnCancel)

        ' =========================================================
        ' BARRA DE PROGRESO (Dock.Bottom)
        ' =========================================================
        _progress = New System.Windows.Forms.ProgressBar()
        _progress.Dock = DockStyle.Bottom
        _progress.Height = 6
        _progress.Style = ProgressBarStyle.Marquee
        _progress.Visible = False

        ' =========================================================
        ' DATA GRID (Anchor: rellena el espacio entre paneles)
        ' =========================================================
        _grid = New DataGridView()
        _grid.Dock = DockStyle.Fill
        _grid.AllowUserToAddRows = False
        _grid.AllowUserToDeleteRows = False
        _grid.AllowUserToResizeRows = False
        _grid.RowHeadersVisible = False
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        _grid.ReadOnly = True
        _grid.BackgroundColor = System.Drawing.Color.White
        _grid.BorderStyle = BorderStyle.Fixed3D
        _grid.ColumnHeadersDefaultCellStyle.Font = New System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold)
        _grid.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(230, 230, 230)
        _grid.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Black
        _grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        _grid.ColumnHeadersHeight = 32
        _grid.EnableHeadersVisualStyles = False
        _grid.DefaultCellStyle.Font = New System.Drawing.Font("Segoe UI", 9)
        _grid.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(100, 160, 220)
        _grid.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Black

        ' Columnas
        Dim colCheck As New DataGridViewCheckBoxColumn()
        colCheck.Name = "Incluir"
        colCheck.HeaderText = "Incluir"
        Dim headerTextSize As System.Drawing.Size = TextRenderer.MeasureText("Incluir", _grid.ColumnHeadersDefaultCellStyle.Font)
        colCheck.Width = headerTextSize.Width + 10
        colCheck.ReadOnly = False
        colCheck.SortMode = DataGridViewColumnSortMode.NotSortable

        Dim colFile As New DataGridViewTextBoxColumn()
        colFile.Name = "Archivo"
        colFile.HeaderText = "Archivo"
        colFile.FillWeight = 220

        Dim colExt As New DataGridViewTextBoxColumn()
        colExt.Name = "Ext"
        colExt.HeaderText = "Ext"
        colExt.Width = 50
        colExt.FillWeight = 50

        Dim colOrigin As New DataGridViewTextBoxColumn()
        colOrigin.Name = "OrigenActual"
        colOrigin.HeaderText = "Origen actual"
        colOrigin.FillWeight = 260

        Dim colDest As New DataGridViewTextBoxColumn()
        colDest.Name = "DestinoPropuesto"
        colDest.HeaderText = "Destino propuesto"
        colDest.FillWeight = 260

        Dim colStatus As New DataGridViewTextBoxColumn()
        colStatus.Name = "Estado"
        colStatus.HeaderText = "Estado"
        colStatus.Width = 140
        colStatus.FillWeight = 140

        _grid.Columns.Add(colCheck)
        _grid.Columns.Add(colFile)
        _grid.Columns.Add(colExt)
        _grid.Columns.Add(colOrigin)
        _grid.Columns.Add(colDest)
        _grid.Columns.Add(colStatus)

        ' Contenedor del grid con Anchor (NO Dock.Fill)
        _gridContainer = New Panel()
        _gridContainer.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        _gridContainer.Controls.Add(_grid)

        AddHandler _grid.CellDoubleClick, AddressOf Grid_CellDoubleClick

        ' =========================================================
        ' ORDEN DE CONTROLES: agregar en orden de Z (de abajo hacia arriba)
        ' =========================================================
        ' En WinForms, el último agregado queda arriba en Z-order.
        ' Queremos que el grid esté al fondo y los paneles arriba.
        Me.Controls.Add(_gridContainer)   ' Fondo
        Me.Controls.Add(_topPanel)        ' Arriba
        Me.Controls.Add(_bottomPanel)     ' Abajo
        Me.Controls.Add(_progress)        ' Abajo (más arriba en Z que bottomPanel)

        ' Calcular posición inicial del grid
        AjustarLayoutGrid()

        ' Recalcular cuando se redimensione
        AddHandler Me.Resize, Sub(s, e) AjustarLayoutGrid()

        UpdateExcludedButton()
        LoadGrid()
    End Sub

    ''' <summary>
    ''' Ajusta la posición y tamaño del contenedor del grid para que quede
    ''' entre el panel superior y los paneles inferiores.
    ''' </summary>
    Private Sub AjustarLayoutGrid()
        Dim topH As Integer = _topPanel.Height
        Dim bottomH As Integer = _bottomPanel.Height + _progress.Height
        _gridContainer.Location = New System.Drawing.Point(0, topH)
        _gridContainer.Size = New System.Drawing.Size(Me.ClientSize.Width, Me.ClientSize.Height - topH - bottomH)
    End Sub

    Private Sub LoadGrid()
        _grid.Rows.Clear()
        Dim itemsToShow = GetVisibleItems()

        For Each item In itemsToShow
            Dim rowIndex = _grid.Rows.Add()
            Dim row = _grid.Rows(rowIndex)

            If item.IsExcluded Then
                row.Cells("Incluir").Value = False
                row.Cells("Incluir").ReadOnly = True
            Else
                row.Cells("Incluir").Value = True
                row.Cells("Incluir").ReadOnly = False
            End If

            row.Cells("Archivo").Value = item.FileName
            row.Cells("Ext").Value = item.Extension.ToLower()
            row.Cells("OrigenActual").Value = GetRelativePath(item.SourcePath)

            If String.IsNullOrEmpty(item.FinalDestination) Then
                row.Cells("DestinoPropuesto").Value = "SIN ASIGNAR"
                row.Cells("DestinoPropuesto").Style.ForeColor = System.Drawing.Color.Red
            Else
                row.Cells("DestinoPropuesto").Value = GetRelativePath(item.FinalDestination)
                row.Cells("DestinoPropuesto").Style.ForeColor = System.Drawing.Color.Black
            End If

            If item.IsExcluded Then
                row.Cells("Estado").Value = "Excluido"
                row.Cells("Estado").Style.ForeColor = System.Drawing.Color.Gray
            ElseIf Not item.IsClassified OrElse String.IsNullOrEmpty(item.TargetFolder) Then
                row.Cells("Estado").Value = "Sin asignar"
                row.Cells("Estado").Style.ForeColor = System.Drawing.Color.Red
            Else
                row.Cells("Estado").Value = "Moviendo a " & item.TargetFolder
                row.Cells("Estado").Style.ForeColor = System.Drawing.Color.Black
            End If

            row.Tag = item

            If item.IsExcluded Then
                row.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(245, 245, 245)
                row.DefaultCellStyle.ForeColor = System.Drawing.Color.Gray
            Else
                Dim bgColor = GetRowColor(item.Extension, item.IsClassified)
                row.DefaultCellStyle.BackColor = bgColor
                row.DefaultCellStyle.ForeColor = System.Drawing.Color.Black
            End If
        Next
    End Sub

    Private Sub Grid_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        If _grid.Columns(e.ColumnIndex).Name <> "DestinoPropuesto" Then Return

        Dim row As DataGridViewRow = _grid.Rows(e.RowIndex)
        Dim item As FileMoveItem = TryCast(row.Tag, FileMoveItem)
        If item Is Nothing OrElse item.IsExcluded Then Return

        Dim initialPath As String = _rootPath
        If Not String.IsNullOrEmpty(item.FinalDestination) Then
            initialPath = IO.Path.GetDirectoryName(item.FinalDestination)
        End If

        Dim picker As New FolderPickerDialog(_rootPath, initialPath)
        If picker.ShowDialog(Me) = DialogResult.OK Then
            Dim selectedFolder As String = picker.SelectedPath
            If String.IsNullOrEmpty(selectedFolder) Then Return

            If Not selectedFolder.StartsWith(_rootPath, StringComparison.OrdinalIgnoreCase) Then
                MessageBox.Show("La carpeta seleccionada debe estar dentro del proyecto.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim newDest As String = IO.Path.Combine(selectedFolder, item.FileName)
            item.FinalDestination = newDest
            item.IsClassified = True

            Dim relPath As String = newDest.Substring(_rootPath.Length).TrimStart("\"c)
            If relPath.Contains("\") Then
                item.TargetFolder = relPath.Split("\"c)(0)
            Else
                item.TargetFolder = relPath
            End If

            row.Cells("DestinoPropuesto").Value = GetRelativePath(newDest)
            row.Cells("DestinoPropuesto").Style.ForeColor = System.Drawing.Color.Black
            row.Cells("Estado").Value = "Moviendo a " & item.TargetFolder
            row.Cells("Estado").Style.ForeColor = System.Drawing.Color.Black

            Dim bgColor As System.Drawing.Color = GetRowColor(item.Extension, item.IsClassified)
            row.DefaultCellStyle.BackColor = bgColor
            row.DefaultCellStyle.ForeColor = System.Drawing.Color.Black
        End If
    End Sub

    Private Function GetRowColor(ext As String, isClassified As Boolean) As System.Drawing.Color
        If Not isClassified Then
            Return System.Drawing.Color.FromArgb(255, 210, 210)
        End If

        Select Case ext.ToLower()
            Case ".ipt" : Return System.Drawing.Color.FromArgb(210, 235, 255)
            Case ".iam" : Return System.Drawing.Color.FromArgb(170, 215, 255)
            Case ".ipn" : Return System.Drawing.Color.FromArgb(210, 255, 210)
            Case ".idw" : Return System.Drawing.Color.FromArgb(170, 235, 170)
            Case ".dwg", ".dxf" : Return System.Drawing.Color.FromArgb(255, 250, 205)
            Case ".pdf" : Return System.Drawing.Color.FromArgb(255, 228, 225)
            Case ".xlsx", ".xls", ".csv" : Return System.Drawing.Color.FromArgb(230, 230, 250)
            Case ".docx", ".doc" : Return System.Drawing.Color.FromArgb(240, 248, 255)
            Case ".jpg", ".jpeg", ".png", ".bmp", ".gif" : Return System.Drawing.Color.FromArgb(255, 240, 245)
            Case ".lnk" : Return System.Drawing.Color.FromArgb(220, 220, 220)
            Case ".txt", ".log" : Return System.Drawing.Color.FromArgb(250, 250, 250)
            Case Else : Return System.Drawing.Color.White
        End Select
    End Function

    Private Function GetVisibleItems() As List(Of FileMoveItem)
        Dim result As New List(Of FileMoveItem)
        For Each item In _allItems
            If Not item.IsExcluded Then result.Add(item)
        Next
        If _showExcluded Then
            For Each item In _allItems
                If item.IsExcluded Then result.Add(item)
            Next
        End If
        If _sortColumn >= 0 Then
            result.Sort(New FileMoveItemComparer(_sortColumn, _sortAscending, _rootPath))
        End If
        Return result
    End Function

    Private Sub ToggleExcluded_Click(sender As Object, e As EventArgs)
        _showExcluded = Not _showExcluded
        UpdateExcludedButton()
        LoadGrid()
    End Sub

    Private Sub UpdateExcludedButton()
        Dim excludedCount As Integer = 0
        For Each i In _allItems
            If i.IsExcluded Then excludedCount += 1
        Next
        Dim arrow = If(_showExcluded, "▼", "▶")
        _btnToggleExcluded.Text = $" {arrow} Excluidos ({excludedCount})"
        _btnToggleExcluded.ForeColor = If(_showExcluded, System.Drawing.Color.DarkRed, System.Drawing.Color.Gray)
    End Sub

    Private Function GetRelativePath(fullPath As String) As String
        If String.IsNullOrEmpty(fullPath) Then Return ""
        If fullPath.StartsWith(_rootPath, StringComparison.OrdinalIgnoreCase) Then
            Return fullPath.Substring(_rootPath.Length).TrimStart("\"c)
        End If
        Return fullPath
    End Function

    ' =========================================================
    ' BOTÓN ACEPTAR: EJECUTA MOVIMIENTOS + LIMPIEZA
    ' =========================================================
    Private Sub BtnOk_Click(sender As Object, e As EventArgs)
        Dim acceptedMoves As New List(Of FileMoveItem)
        For Each row As DataGridViewRow In _grid.Rows
            Dim item = TryCast(row.Tag, FileMoveItem)
            If item Is Nothing Then Continue For
            If item.IsExcluded Then Continue For
            Dim checkCell = TryCast(row.Cells("Incluir"), DataGridViewCheckBoxCell)
            If checkCell IsNot Nothing AndAlso checkCell.Value = True Then
                acceptedMoves.Add(item)
            End If
        Next

        If acceptedMoves.Count = 0 Then
            MessageBox.Show("No hay archivos seleccionados para mover.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim res = MessageBox.Show($"Se van a organizar {acceptedMoves.Count} archivo(s). ¿Continuar?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res <> DialogResult.Yes Then Return

        Me.Enabled = False
        _progress.Visible = True
        _progress.Style = ProgressBarStyle.Marquee
        System.Windows.Forms.Application.DoEvents()

        Dim archivosProcesados As Integer = 0
        Dim archivosIgnorados As Integer = 0
        Dim archivosError As Integer = 0

        ' --- FASE 2: EJECUTAR MOVIMIENTOS ---
        For Each item In acceptedMoves
            Dim file = item.SourcePath
            Dim destino = item.FinalDestination
            Dim nombreArchivo = item.FileName

            If item.IsAlreadyInPlace Then
                archivosIgnorados += 1
                Continue For
            End If

            If String.IsNullOrEmpty(destino) Then
                archivosIgnorados += 1
                Continue For
            End If

            CerrarDocumentoSiEstaAbierto(file)
            CerrarDocumentoSiEstaAbierto(destino)

            Try
                Dim destDir As String = IO.Path.GetDirectoryName(destino)
                If Not IO.Directory.Exists(destDir) Then
                    IO.Directory.CreateDirectory(destDir)
                End If

                IO.File.Copy(file, destino, True)
                IO.File.Delete(file)
                archivosProcesados += 1
                item.Status = "Movido"
            Catch ex As Exception
                archivosError += 1
                item.Status = "Error"
            End Try
        Next

        ' Contar ignorados adicionales
        For Each i In _allItems
            If i.IsExcluded OrElse i.IsAlreadyInPlace OrElse i.Status = "Ignorado" Then
                archivosIgnorados += 1
            End If
        Next

        ' --- FASE 3: ELIMINAR OldVersions (Solo si NO está excluida en Opciones) ---
        Dim oldVersionsEliminadas As Integer = 0
        If Not AddinSettings.Current.ExcluirOldVersions Then
            oldVersionsEliminadas = EliminarOldVersions(_rootPath)
        End If

        ' --- FASE 4: ELIMINAR CARPETAS VACÍAS ---
        Dim carpetasVaciasEliminadas As Integer = 0
        Dim pasada As Integer = 0
        Do
            Dim eliminadasEnEstaPasada = LimpiarCarpetasVacias(_rootPath)
            If eliminadasEnEstaPasada = 0 Then Exit Do
            carpetasVaciasEliminadas += eliminadasEnEstaPasada
            pasada += 1
        Loop While pasada < 10

        _progress.Visible = False
        Me.Enabled = True

        Dim msgResumen As String =
            "Organización completada:" & vbCrLf &
            "Procesados: " & archivosProcesados & vbCrLf &
            "Ignorados: " & archivosIgnorados & vbCrLf &
            "Errores: " & archivosError & vbCrLf & vbCrLf

        If AddinSettings.Current.ExcluirOldVersions Then
            msgResumen &= "Carpetas 'OldVersions': Preservadas (excluidas del borrado)" & vbCrLf
        Else
            msgResumen &= "OldVersions eliminadas: " & oldVersionsEliminadas & vbCrLf
        End If

        msgResumen &= "Carpetas vacías eliminadas: " & carpetasVaciasEliminadas

        MessageBox.Show(
            msgResumen,
            "Organizar Proyecto",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information)

        Me.DialogResult = DialogResult.OK
    End Sub

    ' =========================================================
    ' HELPERS DE EJECUCIÓN
    ' =========================================================
    Private Sub CerrarDocumentoSiEstaAbierto(filePath As String)
        Try
            For Each doc As Document In _invApp.Documents
                If doc.FullFileName.Equals(filePath, StringComparison.OrdinalIgnoreCase) Then
                    Try
                        doc.Close(True)
                    Catch
                    End Try
                    Exit For
                End If
            Next
        Catch
        End Try
    End Sub

    Private Function EliminarOldVersions(rootPath As String) As Integer
        Dim count As Integer = 0
        Try
            If Not IO.Directory.Exists(rootPath) Then Return 0
            For Each subDir As String In IO.Directory.GetDirectories(rootPath)
                count += EliminarOldVersions(subDir)
            Next
            Dim folderName As String = IO.Path.GetFileName(rootPath)
            If folderName.Equals("OldVersions", StringComparison.OrdinalIgnoreCase) Then
                Try
                    IO.Directory.Delete(rootPath, True)
                    count += 1
                Catch
                End Try
            End If
        Catch
        End Try
        Return count
    End Function

    Private Function LimpiarCarpetasVacias(rootPath As String) As Integer
        Dim count As Integer = 0
        Try
            If Not IO.Directory.Exists(rootPath) Then Return 0
            For Each subDir As String In IO.Directory.GetDirectories(rootPath)
                count += LimpiarCarpetasVacias(subDir)
            Next

            Dim folderName As String = IO.Path.GetFileName(rootPath)

            Dim esCarpetaBase As Boolean = False
            For Each cb As String In _carpetasProyecto
                If folderName.Equals(cb, StringComparison.OrdinalIgnoreCase) Then
                    esCarpetaBase = True
                    Exit For
                End If
            Next

            Dim esRaizProyecto As Boolean = String.Equals(rootPath, _rootPath, StringComparison.OrdinalIgnoreCase)

            If esCarpetaBase Or esRaizProyecto Then Return count

            ' Preservar OldVersions si la opción de exclusión está activada
            If AddinSettings.Current.ExcluirOldVersions AndAlso folderName.Equals("OldVersions", StringComparison.OrdinalIgnoreCase) Then
                Return count
            End If

            Dim tieneArchivos As Boolean = IO.Directory.GetFiles(rootPath).Length > 0
            Dim tieneSubcarpetas As Boolean = IO.Directory.GetDirectories(rootPath).Length > 0

            If Not tieneArchivos AndAlso Not tieneSubcarpetas Then
                Try
                    IO.Directory.Delete(rootPath)
                    count += 1
                Catch
                End Try
            End If
        Catch
        End Try
        Return count
    End Function

    ' =========================================================
    ' ORDENAMIENTO
    ' =========================================================
    Private Sub _grid_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles _grid.ColumnHeaderMouseClick
        If e.ColumnIndex = 0 Then Return
        If _sortColumn = e.ColumnIndex Then
            _sortAscending = Not _sortAscending
        Else
            _sortColumn = e.ColumnIndex
            _sortAscending = True
        End If
        For Each col As DataGridViewColumn In _grid.Columns
            Dim clean As String = col.HeaderText.Replace(" ▲", "").Replace(" ▼", "")
            col.HeaderText = clean
        Next
        Dim activeCol = _grid.Columns(_sortColumn)
        activeCol.HeaderText &= If(_sortAscending, " ▲", " ▼")
        LoadGrid()
    End Sub

    Private Class FileMoveItemComparer
        Implements IComparer(Of FileMoveItem)

        Private _columnIndex As Integer
        Private _ascending As Boolean
        Private _rootPath As String

        Public Sub New(columnIndex As Integer, ascending As Boolean, rootPath As String)
            _columnIndex = columnIndex
            _ascending = ascending
            _rootPath = rootPath
        End Sub

        Public Function Compare(x As FileMoveItem, y As FileMoveItem) As Integer Implements IComparer(Of FileMoveItem).Compare
            Dim result As Integer = 0
            Dim sx, sy As String

            Select Case _columnIndex
                Case 1
                    result = String.Compare(x.FileName, y.FileName, StringComparison.OrdinalIgnoreCase)
                Case 2
                    result = String.Compare(x.Extension, y.Extension, StringComparison.OrdinalIgnoreCase)
                Case 3
                    result = String.Compare(x.SourcePath, y.SourcePath, StringComparison.OrdinalIgnoreCase)
                Case 4
                    result = String.Compare(x.FinalDestination, y.FinalDestination, StringComparison.OrdinalIgnoreCase)
                Case 5
                    sx = If(x.IsExcluded, "Excluido", If(Not x.IsClassified OrElse String.IsNullOrEmpty(x.TargetFolder), "Sin asignar", "Moviendo a " & x.TargetFolder))
                    sy = If(y.IsExcluded, "Excluido", If(Not y.IsClassified OrElse String.IsNullOrEmpty(y.TargetFolder), "Sin asignar", "Moviendo a " & y.TargetFolder))
                    result = String.Compare(sx, sy, StringComparison.OrdinalIgnoreCase)
                Case Else
                    result = 0
            End Select

            If Not _ascending Then result = -result
            Return result
        End Function
    End Class

End Class