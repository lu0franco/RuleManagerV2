Imports System
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Linq
Imports System.Runtime.InteropServices
Imports Inventor
''' <summary>
''' Diálogo compacto de asignación COM con indicador recursivo vía StateImageList.
''' Jerarquía visual: [StateImage toggle] [Image tipo] [Texto]
''' </summary>
Public Class ComercialAssignmentDialog
    Inherits Form

    <DllImport("user32.dll")>
    Private Shared Function EnableWindow(hWnd As IntPtr, bEnable As Boolean) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function

    Public Sub HabilitarVentanaInventor()
        If _app IsNot Nothing Then
            Try
                Dim hwnd As New IntPtr(_app.MainFrameHWND)
                If hwnd <> IntPtr.Zero Then
                    EnableWindow(hwnd, True)
                End If
            Catch
            End Try
        End If
    End Sub

    Private _app As Inventor.Application
    Private _treeData As List(Of NodoComercial)
    Private _result As DialogResult = DialogResult.Cancel
    Private _highlightSet As HighlightSet = Nothing

    ' Controles
    Private tvArbol As TreeView
    Private lblResumen As Label
    Private btnAplicar As Button
    Private btnCancelar As Button
    Private btnRefrescar As Button
    Private btnSeleccionarModelo As Button
    Private btnDesmarcarModelo As Button
    Private btnHighlightCom As Button
    Private chkSoloConCambios As CheckBox
    Private txtBuscar As System.Windows.Forms.TextBox
    Private btnExpandir As Button
    Private btnColapsar As Button
    Private btnMarcarTodos As Button
    Private btnDesmarcarTodos As Button

    ' ImageLists
    Private _imageList As ImageList        ' Iconos de tipo (asm, ipt, cc, broken)
    Private _stateImageList As ImageList   ' Toggle visual COM (empty, allok, partial)

    Public ReadOnly Property ShouldApply As Boolean
        Get
            Return _result = DialogResult.OK
        End Get
    End Property

    Public Sub New(app As Inventor.Application)
        _app = app
        InitializeComponent()
        CargarDatos()
    End Sub

    ' =========================================================
    ' INICIALIZACION COMPACTA LATERAL
    ' =========================================================
    Private Sub InitializeComponent()
        Me.Text = "Asignación de Comerciales (COM) - Inspector Jerárquico"
        Me.Size = New System.Drawing.Size(780, 720)
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New System.Drawing.Point(50, 50)
        Me.MinimumSize = New System.Drawing.Size(580, 450)
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.TopMost = True

        AddHandler Me.Shown, Sub(s, e) HabilitarVentanaInventor()
        AddHandler Me.Activated, Sub(s, e) HabilitarVentanaInventor()
        AddHandler Me.FormClosing, Sub(s, e)
                                       ClearHighlight()
                                       HabilitarVentanaInventor()
                                   End Sub

        ' --- Panel izquierdo: TreeView con barra superior de búsqueda y acciones ---
        Dim panelArbol As New Panel()
        panelArbol.Dock = DockStyle.Fill

        Dim pnlArbolHeader As New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 62,
            .BackColor = System.Drawing.Color.FromArgb(241, 245, 249),
            .Padding = New Padding(6, 6, 6, 6)
        }

        ' Fila 1: Búsqueda rápida y botones de expansión
        txtBuscar = New System.Windows.Forms.TextBox() With {
            .Location = New System.Drawing.Point(6, 6),
            .Size = New System.Drawing.Size(260, 23),
            .Font = New Font("Segoe UI", 9.0F)
        }
        Dim tipBuscar As New ToolTip()
        tipBuscar.SetToolTip(txtBuscar, "Escriba texto para filtrar o resaltar componentes por nombre en el árbol.")
        AddHandler txtBuscar.TextChanged, AddressOf TxtBuscar_TextChanged
        pnlArbolHeader.Controls.Add(txtBuscar)

        Dim btnLimpiar As New Button() With {
            .Text = "✕",
            .Location = New System.Drawing.Point(268, 5),
            .Size = New System.Drawing.Size(26, 25),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 7.5F, FontStyle.Bold),
            .Cursor = Cursors.Hand,
            .BackColor = System.Drawing.Color.White
        }
        btnLimpiar.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(203, 213, 225)
        AddHandler btnLimpiar.Click, Sub() txtBuscar.Text = ""
        pnlArbolHeader.Controls.Add(btnLimpiar)

        btnExpandir = New Button() With {
            .Text = "▾ Todo",
            .Location = New System.Drawing.Point(300, 5),
            .Size = New System.Drawing.Size(62, 25),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 8.0F),
            .BackColor = System.Drawing.Color.White,
            .Cursor = Cursors.Hand
        }
        btnExpandir.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(203, 213, 225)
        Dim tipExp As New ToolTip()
        tipExp.SetToolTip(btnExpandir, "Expandir todos los niveles del árbol.")
        AddHandler btnExpandir.Click, Sub() tvArbol.ExpandAll()
        pnlArbolHeader.Controls.Add(btnExpandir)

        btnColapsar = New Button() With {
            .Text = "▸ Colapsar",
            .Location = New System.Drawing.Point(366, 5),
            .Size = New System.Drawing.Size(74, 25),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 8.0F),
            .BackColor = System.Drawing.Color.White,
            .Cursor = Cursors.Hand
        }
        btnColapsar.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(203, 213, 225)
        Dim tipCol As New ToolTip()
        tipCol.SetToolTip(btnColapsar, "Colapsar ramas del árbol.")
        AddHandler btnColapsar.Click, Sub()
                                          tvArbol.CollapseAll()
                                          If tvArbol.Nodes.Count > 0 Then tvArbol.Nodes(0).Expand()
                                      End Sub
        pnlArbolHeader.Controls.Add(btnColapsar)

        btnMarcarTodos = New Button() With {
            .Text = "☑ Marcar",
            .Location = New System.Drawing.Point(446, 5),
            .Size = New System.Drawing.Size(72, 25),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(16, 149, 102),
            .BackColor = System.Drawing.Color.White,
            .Cursor = Cursors.Hand
        }
        btnMarcarTodos.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(167, 243, 208)
        Dim tipMarcar As New ToolTip()
        tipMarcar.SetToolTip(btnMarcarTodos, "Marcar todas las piezas visibles como comerciales.")
        AddHandler btnMarcarTodos.Click, Sub() MarcarTodosLosVisibles(True)
        pnlArbolHeader.Controls.Add(btnMarcarTodos)

        btnDesmarcarTodos = New Button() With {
            .Text = "☐ Desmarcar",
            .Location = New System.Drawing.Point(522, 5),
            .Size = New System.Drawing.Size(84, 25),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(194, 65, 12),
            .BackColor = System.Drawing.Color.White,
            .Cursor = Cursors.Hand
        }
        btnDesmarcarTodos.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(254, 215, 170)
        Dim tipDesmTodos As New ToolTip()
        tipDesmTodos.SetToolTip(btnDesmarcarTodos, "Quitar la marca de comerciales a todas las piezas.")
        AddHandler btnDesmarcarTodos.Click, Sub() MarcarTodosLosVisibles(False)
        pnlArbolHeader.Controls.Add(btnDesmarcarTodos)

        ' Fila 2: Leyenda explicativa
        Dim lblArbol As New Label() With {
            .Text = "[☑]=Tiene COM  [▣]=Parcial  [□]=Sin COM  •  Click casilla = Toggle  •  Doble click = Zoom 3D",
            .Location = New System.Drawing.Point(8, 36),
            .Size = New System.Drawing.Size(595, 20),
            .Font = New Font("Segoe UI", 8.25F),
            .ForeColor = System.Drawing.Color.FromArgb(71, 85, 105)
        }
        pnlArbolHeader.Controls.Add(lblArbol)

        ' Crear ImageLists
        _imageList = CrearImageList()
        _stateImageList = CrearStateImageList()

        tvArbol = New TreeView()
        tvArbol.Dock = DockStyle.Fill
        tvArbol.CheckBoxes = False
        tvArbol.ImageList = _imageList
        tvArbol.StateImageList = _stateImageList
        tvArbol.FullRowSelect = True
        tvArbol.Font = New Font("Segoe UI", 8.75F)

        AddHandler tvArbol.NodeMouseClick, AddressOf TvArbol_NodeMouseClick
        AddHandler tvArbol.NodeMouseDoubleClick, AddressOf TvArbol_NodeMouseDoubleClick
        AddHandler tvArbol.MouseClick, AddressOf TvArbol_MouseClick

        panelArbol.Controls.Add(tvArbol)
        panelArbol.Controls.Add(pnlArbolHeader)

        ' --- Panel derecho: Botones de Acción ---
        Dim panelBotones As New Panel()
        panelBotones.Dock = DockStyle.Right
        panelBotones.Width = 165
        panelBotones.BackColor = System.Drawing.Color.FromArgb(248, 250, 252)
        panelBotones.Padding = New Padding(8, 10, 8, 10)

        Dim y As Integer = 10

        Dim lblTitulo As New Label()
        lblTitulo.Text = "Acciones"
        lblTitulo.Location = New System.Drawing.Point(8, y)
        lblTitulo.Size = New System.Drawing.Size(148, 20)
        lblTitulo.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        lblTitulo.ForeColor = System.Drawing.Color.FromArgb(15, 23, 42)
        panelBotones.Controls.Add(lblTitulo)
        y += 28

        btnSeleccionarModelo = CrearBotonLateral("Seleccionar del modelo", y, System.Drawing.Color.FromArgb(16, 149, 102))
        Dim tipSel As New ToolTip()
        tipSel.SetToolTip(btnSeleccionarModelo, "Selección interactiva continua en el modelo 3D (Alternar / Marcar / Desmarcar).")
        AddHandler btnSeleccionarModelo.Click, AddressOf BtnSeleccionarModelo_Click
        panelBotones.Controls.Add(btnSeleccionarModelo)
        y += 40

        btnDesmarcarModelo = CrearBotonLateral("Desmarcar del modelo", y, System.Drawing.Color.FromArgb(217, 83, 79))
        Dim tipDesm As New ToolTip()
        tipDesm.SetToolTip(btnDesmarcarModelo, "Haga clic en las piezas en el modelo 3D que desea desmarcar (quitar COM).")
        AddHandler btnDesmarcarModelo.Click, AddressOf BtnDesmarcarModelo_Click
        panelBotones.Controls.Add(btnDesmarcarModelo)
        y += 40

        btnHighlightCom = CrearBotonLateral("Resaltar COM en 3D", y, System.Drawing.Color.FromArgb(220, 38, 38))
        Dim tipHl As New ToolTip()
        tipHl.SetToolTip(btnHighlightCom, "Resalta en rojo todas las piezas comerciales en el modelo 3D.")
        AddHandler btnHighlightCom.Click, AddressOf BtnHighlightCom_Click
        panelBotones.Controls.Add(btnHighlightCom)
        y += 46

        Dim sep1 As New Label()
        sep1.Location = New System.Drawing.Point(8, y)
        sep1.Size = New System.Drawing.Size(148, 1)
        sep1.BackColor = System.Drawing.Color.FromArgb(226, 232, 240)
        panelBotones.Controls.Add(sep1)
        y += 10

        chkSoloConCambios = New CheckBox()
        chkSoloConCambios.Text = "Solo con cambios"
        chkSoloConCambios.Location = New System.Drawing.Point(8, y)
        chkSoloConCambios.Size = New System.Drawing.Size(148, 22)
        chkSoloConCambios.Font = New Font("Segoe UI", 8.25F)
        chkSoloConCambios.ForeColor = System.Drawing.Color.FromArgb(30, 41, 59)
        chkSoloConCambios.AutoSize = False
        AddHandler chkSoloConCambios.CheckedChanged, AddressOf ChkSoloConCambios_CheckedChanged
        panelBotones.Controls.Add(chkSoloConCambios)
        y += 28

        btnRefrescar = CrearBotonLateral("Refrescar", y, System.Drawing.Color.FromArgb(100, 116, 139))
        Dim tipRef As New ToolTip()
        tipRef.SetToolTip(btnRefrescar, "Recarga la estructura completa leyendo las iProperties reales de los archivos.")
        AddHandler btnRefrescar.Click, AddressOf BtnRefrescar_Click
        panelBotones.Controls.Add(btnRefrescar)
        y += 44

        Dim sep2 As New Label()
        sep2.Location = New System.Drawing.Point(8, y)
        sep2.Size = New System.Drawing.Size(148, 1)
        sep2.BackColor = System.Drawing.Color.FromArgb(226, 232, 240)
        panelBotones.Controls.Add(sep2)
        y += 10

        btnAplicar = CrearBotonLateral("Aplicar cambios", y, System.Drawing.Color.FromArgb(2, 132, 199))
        Dim tipAp As New ToolTip()
        tipAp.SetToolTip(btnAplicar, "Escribe la propiedad COM en todos los archivos modificados y actualiza el ensamble.")
        AddHandler btnAplicar.Click, AddressOf BtnAplicar_Click
        panelBotones.Controls.Add(btnAplicar)
        y += 40

        btnCancelar = CrearBotonLateral("Cancelar", y, System.Drawing.Color.FromArgb(100, 116, 139))
        AddHandler btnCancelar.Click, AddressOf BtnCancelar_Click
        panelBotones.Controls.Add(btnCancelar)
        y += 44

        Dim sep3 As New Label()
        sep3.Location = New System.Drawing.Point(8, y)
        sep3.Size = New System.Drawing.Size(148, 1)
        sep3.BackColor = System.Drawing.Color.FromArgb(226, 232, 240)
        panelBotones.Controls.Add(sep3)
        y += 10

        lblResumen = New Label()
        lblResumen.Text = "Listo"
        lblResumen.Location = New System.Drawing.Point(8, y)
        lblResumen.Size = New System.Drawing.Size(148, 90)
        lblResumen.Font = New Font("Segoe UI", 8.25F)
        lblResumen.ForeColor = System.Drawing.Color.FromArgb(71, 85, 105)
        lblResumen.AutoSize = False
        panelBotones.Controls.Add(lblResumen)

        Me.Controls.Add(panelArbol)
        Me.Controls.Add(panelBotones)

        Me.AcceptButton = btnAplicar
        Me.CancelButton = btnCancelar
    End Sub

    Private Function CrearBotonLateral(texto As String, y As Integer, colorFondo As System.Drawing.Color) As Button
        Dim btn As New Button()
        btn.Text = texto
        btn.Location = New System.Drawing.Point(8, y)
        btn.Size = New System.Drawing.Size(148, 34)
        btn.BackColor = colorFondo
        btn.ForeColor = System.Drawing.Color.White
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 0
        btn.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        btn.TextAlign = ContentAlignment.MiddleCenter
        btn.Cursor = Cursors.Hand
        Return btn
    End Function

    Private Function CrearImageList() As ImageList
        Dim il As New ImageList()
        il.ImageSize = New System.Drawing.Size(16, 16)
        il.ColorDepth = ColorDepth.Depth32Bit

        il.Images.Add("asm", CrearCirculo(System.Drawing.Color.Blue))
        il.Images.Add("ipt", CrearCirculo(System.Drawing.Color.Orange))
        il.Images.Add("cc", CrearCirculo(System.Drawing.Color.Gray))
        il.Images.Add("broken", CrearCirculo(System.Drawing.Color.Red))

        Return il
    End Function

    Private Function CrearCirculo(color As System.Drawing.Color) As Bitmap
        Dim bmp As New Bitmap(16, 16)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            g.Clear(System.Drawing.Color.Transparent)
            g.FillEllipse(New SolidBrush(color), 2, 2, 12, 12)
        End Using
        Return bmp
    End Function

    Private Function CrearStateImageList() As ImageList
        Dim il As New ImageList()
        il.ImageSize = New System.Drawing.Size(16, 16)
        il.ColorDepth = ColorDepth.Depth32Bit

        il.Images.Add("empty", CrearEstadoVacio())
        il.Images.Add("allok", CrearEstadoTodoOK())
        il.Images.Add("partial", CrearEstadoParcial())

        Return il
    End Function

    Private Function CrearEstadoVacio() As Bitmap
        Dim bmp As New Bitmap(16, 16)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            g.Clear(System.Drawing.Color.Transparent)
            Using pen As New Pen(System.Drawing.Color.FromArgb(160, 160, 160), 1.5)
                g.DrawRectangle(pen, 3, 3, 10, 10)
            End Using
        End Using
        Return bmp
    End Function

    Private Function CrearEstadoTodoOK() As Bitmap
        Dim bmp As New Bitmap(16, 16)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            g.Clear(System.Drawing.Color.Transparent)
            g.FillRectangle(New SolidBrush(System.Drawing.Color.FromArgb(0, 120, 212)), 3, 3, 10, 10)
            Using pen As New Pen(System.Drawing.Color.White, 2)
                pen.EndCap = Drawing2D.LineCap.Round
                pen.StartCap = Drawing2D.LineCap.Round
                g.DrawLine(pen, 5, 8, 7, 10)
                g.DrawLine(pen, 7, 10, 11, 6)
            End Using
        End Using
        Return bmp
    End Function

    Private Function CrearEstadoParcial() As Bitmap
        Dim bmp As New Bitmap(16, 16)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            g.Clear(System.Drawing.Color.Transparent)
            g.FillRectangle(New SolidBrush(System.Drawing.Color.FromArgb(0, 120, 212)), 3, 3, 10, 10)
            Using pen As New Pen(System.Drawing.Color.White, 2)
                pen.EndCap = Drawing2D.LineCap.Round
                pen.StartCap = Drawing2D.LineCap.Round
                g.DrawLine(pen, 5, 8, 11, 8)
            End Using
        End Using
        Return bmp
    End Function

    ' =========================================================
    ' CARGAR DATOS
    ' =========================================================
    Private Sub CargarDatos()
        tvArbol.Nodes.Clear()
        _treeData = New List(Of NodoComercial)

        Try
            Dim doc As Document = _app.ActiveDocument
            If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("No hay un ensamblaje activo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
            Dim rootNode As New TreeNode("Ensamblaje principal")
            rootNode.ImageKey = "asm"
            rootNode.SelectedImageKey = "asm"
            rootNode.StateImageIndex = 0

            For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                Dim nodo = CrearNodoRecursivo(occ)
                If nodo IsNot Nothing Then
                    rootNode.Nodes.Add(nodo)
                End If
            Next

            CalcularEstadosRecursivosInicial(rootNode)

            tvArbol.Nodes.Add(rootNode)
            rootNode.Expand()

            ActualizarPreview()
            HighlightByComValue()

        Catch ex As Exception
            MessageBox.Show("Error cargando datos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' =========================================================
    ' CREAR NODO RECURSIVO
    ' =========================================================
    Private Function CrearNodoRecursivo(occ As ComponentOccurrence) As TreeNode
        Dim datos As NodoComercial = ExtraerDatos(occ)

        If datos.IsBuloneria Then
            _treeData.Add(datos)
            Return Nothing
        End If

        _treeData.Add(datos)

        Dim nodo As New TreeNode()
        nodo.Text = datos.FileNameShort & If(datos.IsContentCenter, " [CC]", "")
        nodo.Tag = datos
        nodo.StateImageIndex = 0

        If datos.IsBroken Then
            nodo.ImageKey = "broken"
        ElseIf datos.IsContentCenter Then
            nodo.ImageKey = "cc"
        Else
            nodo.ImageKey = If(datos.DefinitionType = DocumentTypeEnum.kAssemblyDocumentObject, "asm", "ipt")
        End If
        nodo.SelectedImageKey = nodo.ImageKey

        If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject AndAlso Not datos.IsBroken Then
            For Each subOcc As ComponentOccurrence In occ.SubOccurrences
                Dim subNodo = CrearNodoRecursivo(subOcc)
                If subNodo IsNot Nothing Then
                    nodo.Nodes.Add(subNodo)
                End If
            Next
        End If

        Return nodo
    End Function

    Private Sub CalcularEstadosRecursivosInicial(nodo As TreeNode)
        If nodo Is Nothing Then Return
        CalcularEstadoRecursivoReal(nodo)
    End Sub

    ' =========================================================
    ' CALCULAR ESTADO RECURSIVO REAL
    ' =========================================================
    Private Function CalcularEstadoRecursivoReal(nodo As TreeNode) As Boolean
        If nodo Is Nothing Then Return False

        Dim datos As NodoComercial = TryCast(nodo.Tag, NodoComercial)

        If datos IsNot Nothing AndAlso (datos.IsBroken OrElse datos.IsBuloneria) Then
            nodo.StateImageIndex = 0
            Return False
        End If

        If datos IsNot Nothing AndAlso datos.IsContentCenter Then
            Dim ccTieneCom As Boolean = (datos.ComActual = "✓")
            nodo.StateImageIndex = If(ccTieneCom, 1, 0)
            Return ccTieneCom
        End If

        Dim tieneComPropio As Boolean = If(datos IsNot Nothing, (datos.ComNuevo = "✓"), False)

        Dim todosHijosConCom As Boolean = True
        Dim algunHijoConCom As Boolean = False
        Dim tieneHijosValidos As Boolean = False

        For Each hijo As TreeNode In nodo.Nodes
            Dim hijoDatos As NodoComercial = TryCast(hijo.Tag, NodoComercial)

            If hijoDatos IsNot Nothing AndAlso (hijoDatos.IsBroken OrElse hijoDatos.IsBuloneria) Then
                Continue For
            End If

            tieneHijosValidos = True
            Dim hijoTieneCom As Boolean = CalcularEstadoRecursivoReal(hijo)

            If hijoTieneCom OrElse hijo.StateImageIndex = 1 OrElse hijo.StateImageIndex = 2 Then
                algunHijoConCom = True
            End If

            If Not hijoTieneCom Then
                todosHijosConCom = False
            End If
        Next

        If tieneComPropio Then
            nodo.StateImageIndex = 1
        ElseIf algunHijoConCom Then
            nodo.StateImageIndex = 2
        Else
            nodo.StateImageIndex = 0
        End If

        Return tieneComPropio OrElse (tieneHijosValidos AndAlso todosHijosConCom)
    End Function

    ' =========================================================
    ' EXTRAER DATOS
    ' =========================================================
    Private Function ExtraerDatos(occ As ComponentOccurrence) As NodoComercial
        Dim datos As New NodoComercial()

        Try
            datos.OccurrenceName = occ.Name

            If occ.Definition Is Nothing OrElse occ.Definition.Document Is Nothing Then
                datos.IsBroken = True
                datos.FileNameShort = occ.Name & " [ROTO]"
                datos.DefinitionType = DocumentTypeEnum.kAssemblyDocumentObject
                Return datos
            End If

            Dim doc As Document = occ.Definition.Document
            datos.FileName = doc.FullDocumentName
            datos.FileNameShort = System.IO.Path.GetFileName(datos.FileName)
            datos.ComActual = LeerComDeDocumento(doc)
            datos.IsContentCenter = (occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject AndAlso occ.Definition.IsContentMember)
            datos.IsBuloneria = EsBuloneria(doc)
            datos.DefinitionType = occ.DefinitionDocumentType

            datos.ComNuevo = datos.ComActual
            datos.TieneCambio = False
            datos.Occurrence = occ

        Catch ex As Exception
            datos.IsBroken = True
            datos.FileNameShort = occ.Name & " [ERROR]"
            datos.DefinitionType = DocumentTypeEnum.kAssemblyDocumentObject
        End Try

        Return datos
    End Function

    Private Function EsBuloneria(doc As Document) As Boolean
        Try
            Dim props As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
            For Each p As Inventor.Property In props
                If p.Name = "BULONERIA" AndAlso p.Value.ToString().Trim() = "✓" Then
                    Return True
                End If
            Next
        Catch
        End Try
        Return False
    End Function

    Private Function LeerComDeDocumento(doc As Document) As String
        Try
            Dim props As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
            For Each p As Inventor.Property In props
                If p.Name = "COM" Then
                    Return p.Value.ToString().Trim()
                End If
            Next
            Return ""
        Catch
            Return ""
        End Try
    End Function

    ' =========================================================
    ' ACTUALIZAR RESUMEN
    ' =========================================================
    Private Sub ActualizarPreview()
        Dim cambios = _treeData.Where(Function(n) n.TieneCambio AndAlso Not n.IsBroken AndAlso Not n.IsBuloneria).ToList()
        Dim comercialesCount As Integer = _treeData.Where(Function(n) n.ComActual = "✓" AndAlso Not n.IsBuloneria).Count()
        Dim ccCount As Integer = _treeData.Where(Function(n) n.IsContentCenter).Count()
        Dim buloneriaCount As Integer = _treeData.Where(Function(n) n.IsBuloneria).Count()
        Dim visiblesCount As Integer = _treeData.Where(Function(n) Not n.IsBuloneria).Count()

        lblResumen.Text = String.Format(
            "Visibles: {0}" & vbCrLf &
            "Cambios: {1}" & vbCrLf &
            "Con COM: {2}" & vbCrLf &
            "CC: {3}" & vbCrLf &
            "Buloneria: {4}",
            visiblesCount, cambios.Count, comercialesCount, ccCount, buloneriaCount)
    End Sub

    ' =========================================================
    ' EVENTOS MOUSE
    ' =========================================================
    Private Sub TvArbol_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs)
        If e.Button <> MouseButtons.Left Then Return

        Dim hit As TreeViewHitTestInfo = tvArbol.HitTest(e.Location)
        If hit.Location = TreeViewHitTestLocations.StateImage Then
            ToggleComNodo(e.Node)
        End If
    End Sub

    Private Sub ToggleComNodo(nodo As TreeNode)
        If nodo Is Nothing OrElse nodo.Tag Is Nothing Then Return

        Dim datos As NodoComercial = CType(nodo.Tag, NodoComercial)

        If datos.IsContentCenter OrElse datos.IsBroken OrElse datos.IsBuloneria Then Return

        Dim nuevoState As Integer
        If nodo.StateImageIndex = 1 Then
            nuevoState = 0
            datos.ComNuevo = ""
        Else
            nuevoState = 1
            datos.ComNuevo = "✓"
        End If

        datos.TieneCambio = (datos.ComActual <> datos.ComNuevo)
        nodo.StateImageIndex = nuevoState

        If datos.Occurrence IsNot Nothing Then
            SyncHighlightWithCheckbox(datos.Occurrence, nuevoState = 1)
        End If

        RecalcularEstadoHaciaArriba(nodo)
        ActualizarPreview()
    End Sub

    Private Sub RecalcularEstadoHaciaArriba(nodo As TreeNode)
        If nodo Is Nothing Then Return

        CalcularEstadoRecursivoReal(nodo)

        If nodo.Parent IsNot Nothing Then
            RecalcularEstadoHaciaArriba(nodo.Parent)
        End If
    End Sub

    Private Sub TvArbol_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs)
        If e.Node.Tag Is Nothing Then Return

        Dim datos As NodoComercial = CType(e.Node.Tag, NodoComercial)
        If datos.Occurrence Is Nothing Then Return

        Try
            ToggleWholeFamily(datos.Occurrence)

            Dim doc As Document = _app.ActiveDocument
            Dim occ As ComponentOccurrence = datos.Occurrence

            doc.SelectSet.Clear()
            doc.SelectSet.Select(occ)

            Dim view As Inventor.View = _app.ActiveView
            If view IsNot Nothing Then
                Dim cam As Camera = view.Camera
                Dim box As Box = occ.RangeBox

                Dim center As Inventor.Point = _app.TransientGeometry.CreatePoint(
                    (box.MinPoint.X + box.MaxPoint.X) / 2.0,
                    (box.MinPoint.Y + box.MaxPoint.Y) / 2.0,
                    (box.MinPoint.Z + box.MaxPoint.Z) / 2.0)

                Dim sizeX As Double = Math.Abs(box.MaxPoint.X - box.MinPoint.X)
                Dim sizeY As Double = Math.Abs(box.MaxPoint.Y - box.MinPoint.Y)
                Dim sizeZ As Double = Math.Abs(box.MaxPoint.Z - box.MinPoint.Z)
                Dim maxSize As Double = Math.Max(sizeX, Math.Max(sizeY, sizeZ))
                Dim distance As Double = Math.Max(maxSize * 2.5, 10.0)

                cam.Target = center
                cam.Eye = _app.TransientGeometry.CreatePoint(
                    center.X + distance,
                    center.Y + distance,
                    center.Z + distance)
                cam.UpVector = _app.TransientGeometry.CreateUnitVector(0, 0, 1)
                cam.Apply()
                view.Update()
            End If

            doc.SelectSet.Clear()
            HabilitarVentanaInventor()
        Catch
        End Try
    End Sub

    Private Sub TvArbol_MouseClick(sender As Object, e As MouseEventArgs)
        If e.Button <> MouseButtons.Right Then Return

        Dim node As TreeNode = tvArbol.GetNodeAt(e.Location)
        If node Is Nothing OrElse node.Tag Is Nothing Then Return

        tvArbol.SelectedNode = node
        Dim menu As New ContextMenuStrip()

        Dim itemMarcarRecursivo As New ToolStripMenuItem("Marcar COM recursivo")
        AddHandler itemMarcarRecursivo.Click, Sub(s, args) MarcarComRecursivo(node, True)
        menu.Items.Add(itemMarcarRecursivo)

        Dim itemLimpiarRecursivo As New ToolStripMenuItem("Limpiar COM recursivo")
        AddHandler itemLimpiarRecursivo.Click, Sub(s, args) MarcarComRecursivo(node, False)
        menu.Items.Add(itemLimpiarRecursivo)

        menu.Items.Add(New ToolStripSeparator())

        Dim itemVer As New ToolStripMenuItem("Ver en modelo")
        AddHandler itemVer.Click, Sub(s, args)
                                      Dim datos As NodoComercial = CType(node.Tag, NodoComercial)
                                      If datos.Occurrence IsNot Nothing Then
                                          TvArbol_NodeMouseDoubleClick(Nothing, New TreeNodeMouseClickEventArgs(node, MouseButtons.Left, 1, 0, 0))
                                      End If
                                  End Sub
        menu.Items.Add(itemVer)

        Dim itemPropiedades As New ToolStripMenuItem("Ver propiedades...")
        AddHandler itemPropiedades.Click, Sub(s, args)
                                              Dim datos As NodoComercial = CType(node.Tag, NodoComercial)
                                              MessageBox.Show(
                                                  "Archivo: " & datos.FileName & vbCrLf &
                                                  "COM Actual: " & If(datos.ComActual = "", "(vacio)", datos.ComActual) & vbCrLf &
                                                  "COM Nuevo: " & If(datos.ComNuevo = "", "(vacio)", datos.ComNuevo) & vbCrLf &
                                                  "ContentCenter: " & datos.IsContentCenter & vbCrLf &
                                                  "Buloneria: " & datos.IsBuloneria,
                                                  "Propiedades de " & datos.FileNameShort,
                                                  MessageBoxButtons.OK, MessageBoxIcon.Information)
                                          End Sub
        menu.Items.Add(itemPropiedades)

        menu.Show(tvArbol, e.Location)
    End Sub

    Private Sub MarcarComRecursivo(nodo As TreeNode, marcar As Boolean)
        If nodo Is Nothing Then Return

        If nodo.Tag IsNot Nothing Then
            Dim datos As NodoComercial = CType(nodo.Tag, NodoComercial)

            If Not datos.IsContentCenter AndAlso Not datos.IsBuloneria AndAlso Not datos.IsBroken Then
                Dim nuevoState As Integer = If(marcar, 1, 0)
                If nodo.StateImageIndex <> nuevoState Then
                    nodo.StateImageIndex = nuevoState
                    datos.ComNuevo = If(marcar, "✓", "")
                    datos.TieneCambio = (datos.ComActual <> datos.ComNuevo)

                    If datos.Occurrence IsNot Nothing Then
                        SyncHighlightWithCheckbox(datos.Occurrence, marcar)
                    End If
                End If
            End If
        End If

        For Each hijo As TreeNode In nodo.Nodes
            MarcarComRecursivo(hijo, marcar)
        Next

        RecalcularEstadoHaciaArriba(nodo)
        ActualizarPreview()
    End Sub

    ' =========================================================
    ' SELECCION EN MODELO 3D (INTERACTIVA, CONTINUA Y CON DESMARCAR)
    ' =========================================================
    Private Sub BtnSeleccionarModelo_Click(sender As Object, e As EventArgs)
        IniciarSeleccionEnModelo(ModoAccionModelo.Alternar)
    End Sub

    Private Sub BtnDesmarcarModelo_Click(sender As Object, e As EventArgs)
        IniciarSeleccionEnModelo(ModoAccionModelo.Desmarcar)
    End Sub

    Private Sub IniciarSeleccionEnModelo(modoInicial As ModoAccionModelo)
        Try
            Dim doc As Document = _app.ActiveDocument
            If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show(Me, "No hay un ensamblaje activo (.iam)", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            doc.SelectSet.Clear()
            Me.Hide() ' Ocultar temporalmente para interactuar con Inventor

            ' Inicializar el asistente interactivo continuo
            Dim helper As New SeleccionModeloHelper(_app, Me, modoInicial)
            helper.IniciarSeleccion()

        Catch ex As Exception
            Me.Show()
            Me.Activate()
            HabilitarVentanaInventor()
            MessageBox.Show(Me, "Error al iniciar selección en modelo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Notificación llamada cuando el usuario presiona ESC o Finalizar en la barra flotante.
    ''' Restaura el diálogo principal y refresca las marcas y números en pantalla.
    ''' </summary>
    Public Sub OnSeleccionModeloFinalizada()
        Me.Show()
        Me.Activate()
        Me.WindowState = FormWindowState.Normal
        HabilitarVentanaInventor()

        RecalcularTodosLosEstados()
        ActualizarPreview()
    End Sub

    ''' <summary>
    ''' Procesa en tiempo real el click del usuario sobre cualquier pieza o sub-ensamblaje del 3D.
    ''' Permite Marcar (+), Desmarcar (-) o Alternar (Toggle), sincronizando todas las instancias de la familia.
    ''' </summary>
    Public Function ProcesarClickModeloEnVivo(occ As ComponentOccurrence, modo As ModoAccionModelo, recursivo As Boolean) As ResultadoClickModelo
        Dim res As New ResultadoClickModelo()
        If occ Is Nothing Then Return res

        Try
            If occ.Definition Is Nothing OrElse occ.Definition.Document Is Nothing Then
                res.Mensaje = "Componente sin documento o roto"
                Return res
            End If

            Dim doc As Document = occ.Definition.Document
            Dim key As String = doc.FullDocumentName
            Dim shortName As String = System.IO.Path.GetFileName(key)

            If EsBuloneria(doc) Then
                res.Mensaje = "Bulonería omitida: " & shortName
                Return res
            End If

            ' Buscar todos los nodos en memoria que correspondan a este archivo
            Dim matchingNodes As List(Of NodoComercial) = _treeData.Where(Function(n) String.Equals(n.FileName, key, StringComparison.OrdinalIgnoreCase)).ToList()
            If matchingNodes.Count = 0 Then
                Dim nuevoDato As NodoComercial = ExtraerDatos(occ)
                _treeData.Add(nuevoDato)
                matchingNodes.Add(nuevoDato)
            End If

            ' Determinar si marcamos o desmarcamos
            Dim marcar As Boolean
            Select Case modo
                Case ModoAccionModelo.Marcar
                    marcar = True
                Case ModoAccionModelo.Desmarcar
                    marcar = False
                Case Else ' Alternar
                    Dim estadoActual As Boolean = (matchingNodes(0).ComNuevo = "✓")
                    marcar = Not estadoActual
            End Select

            ' Actualizar todos los datos en memoria para todas las ocurrencias del mismo archivo
            For Each n In matchingNodes
                n.ComNuevo = If(marcar, "✓", "")
                n.TieneCambio = (n.ComActual <> n.ComNuevo)
            Next

            ' Actualizar visualmente los nodos del TreeView
            ActualizarNodosEnTreePorArchivo(key, marcar)

            ' Actualizar HighlightSet en el visor 3D para toda la familia
            Dim hs As HighlightSet = GetOrCreateHighlightSet()
            If hs IsNot Nothing Then
                Dim asmDoc As AssemblyDocument = TryCast(_app.ActiveDocument, AssemblyDocument)
                If asmDoc IsNot Nothing Then
                    RecorrerToggleHighlight(asmDoc.ComponentDefinition.Occurrences, key, hs, quitar:=(Not marcar))
                    _app.ActiveView.Update()
                End If
            End If

            ' Si es recursivo y es ensamblaje, propagar a sub-ocurrencias
            If recursivo AndAlso occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                For Each subOcc As ComponentOccurrence In occ.SubOccurrences
                    ProcesarClickModeloEnVivo(subOcc, If(marcar, ModoAccionModelo.Marcar, ModoAccionModelo.Desmarcar), True)
                Next
            End If

            RecalcularTodosLosEstados()
            ActualizarPreview()

            res.CantidadAfectada = matchingNodes.Count
            res.EsMarcado = marcar
            res.Mensaje = String.Format("{0}: {1} ({2} inst.)", If(marcar, "✓ Marcado COM", "✕ Desmarcado (Sin COM)"), shortName, matchingNodes.Count)

        Catch ex As Exception
            res.Mensaje = "Error: " & ex.Message
        End Try

        Return res
    End Function

    Private Sub ActualizarNodosEnTreePorArchivo(key As String, marcar As Boolean)
        RecorrerTreeActualizarNodo(tvArbol.Nodes, key, marcar)
    End Sub

    Private Sub RecorrerTreeActualizarNodo(nodes As TreeNodeCollection, key As String, marcar As Boolean)
        For Each node As TreeNode In nodes
            Dim datos As NodoComercial = TryCast(node.Tag, NodoComercial)
            If datos IsNot Nothing AndAlso String.Equals(datos.FileName, key, StringComparison.OrdinalIgnoreCase) Then
                node.StateImageIndex = If(marcar, 1, 0)
            End If
            If node.Nodes.Count > 0 Then
                RecorrerTreeActualizarNodo(node.Nodes, key, marcar)
            End If
        Next
    End Sub

    Private Sub RecalcularTodosLosEstados()
        For Each rootNode As TreeNode In tvArbol.Nodes
            CalcularEstadoRecursivoReal(rootNode)
        Next
    End Sub

    Private Sub MarcarTodosLosVisibles(marcar As Boolean)
        For Each n In _treeData
            If Not n.IsBuloneria AndAlso Not n.IsBroken AndAlso Not n.IsContentCenter Then
                n.ComNuevo = If(marcar, "✓", "")
                n.TieneCambio = (n.ComActual <> n.ComNuevo)
            End If
        Next

        For Each root As TreeNode In tvArbol.Nodes
            MarcarComRecursivo(root, marcar)
        Next

        RecalcularTodosLosEstados()
        ActualizarPreview()
        ClearHighlight()
        HighlightByComValue()
    End Sub

    ' =========================================================
    ' BUSQUEDA Y FILTRADO RAPIDO EN EL ARBOL
    ' =========================================================
    Private Sub TxtBuscar_TextChanged(sender As Object, e As EventArgs)
        Dim filtro As String = If(txtBuscar IsNot Nothing, txtBuscar.Text.Trim().ToLowerInvariant(), "")
        If String.IsNullOrEmpty(filtro) Then
            For Each n As TreeNode In tvArbol.Nodes
                RestaurarColorNodo(n)
            Next
            Return
        End If

        tvArbol.BeginUpdate()
        For Each root As TreeNode In tvArbol.Nodes
            FiltrarNodosPorTexto(root, filtro)
        Next
        tvArbol.EndUpdate()
    End Sub

    Private Function FiltrarNodosPorTexto(nodo As TreeNode, filtro As String) As Boolean
        Dim coincide As Boolean = nodo.Text.ToLowerInvariant().Contains(filtro)
        Dim algunHijoCoincide As Boolean = False

        For Each hijo As TreeNode In nodo.Nodes
            If FiltrarNodosPorTexto(hijo, filtro) Then
                algunHijoCoincide = True
            End If
        Next

        If coincide OrElse algunHijoCoincide Then
            nodo.BackColor = If(coincide, System.Drawing.Color.FromArgb(254, 240, 138), System.Drawing.Color.White)
            nodo.ForeColor = System.Drawing.Color.Black
            If algunHijoCoincide Then nodo.Expand()
            Return True
        Else
            nodo.BackColor = System.Drawing.Color.White
            nodo.ForeColor = System.Drawing.Color.FromArgb(160, 160, 160)
            Return False
        End If
    End Function

    Private Sub RestaurarColorNodo(nodo As TreeNode)
        nodo.BackColor = System.Drawing.Color.White
        nodo.ForeColor = System.Drawing.Color.Black
        For Each hijo As TreeNode In nodo.Nodes
            RestaurarColorNodo(hijo)
        Next
    End Sub

    ' =========================================================
    ' HIGHLIGHTS
    ' =========================================================
    Private Sub BtnHighlightCom_Click(sender As Object, e As EventArgs)
        Try
            ClearHighlight()
            HighlightByComValue()
        Catch ex As Exception
            MessageBox.Show("Error al resaltar: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function GetOrCreateHighlightSet() As HighlightSet
        Dim doc As Document = _app.ActiveDocument
        If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then Return Nothing
        Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
        Try
            _highlightSet = asmDoc.HighlightSets.Item(1)
        Catch
            _highlightSet = asmDoc.HighlightSets.Add()
            _highlightSet.Color = _app.TransientObjects.CreateColor(255, 0, 0)
        End Try
        Return _highlightSet
    End Function

    Public Sub HighlightByComValue()
        Dim hs As HighlightSet = GetOrCreateHighlightSet()
        If hs Is Nothing Then Return
        Dim doc As Document = _app.ActiveDocument
        If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then Return
        Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
        hs.Clear()
        RecorrerComHighlight(asmDoc.ComponentDefinition.Occurrences, hs)
        _app.ActiveView.Update()
        HabilitarVentanaInventor()
    End Sub

    Private Sub RecorrerComHighlight(occs As ComponentOccurrences, hs As HighlightSet)
        For Each occ As ComponentOccurrence In occs
            If occ.Suppressed Then Continue For
            Dim tieneTick As Boolean = False
            Try
                If occ.Definition IsNot Nothing AndAlso occ.Definition.Document IsNot Nothing Then
                    Dim ps As PropertySet = occ.Definition.Document.PropertySets.Item("Inventor User Defined Properties")
                    tieneTick = (CStr(ps.Item("COM").Value).Trim() = "✓")
                End If
            Catch
                tieneTick = False
            End Try

            If tieneTick Then
                Try : hs.AddItem(occ) : Catch : End Try
            End If

            If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                RecorrerComHighlight(occ.SubOccurrences, hs)
            End If
        Next
    End Sub

    Public Sub ToggleWholeFamily(occ As ComponentOccurrence)
        Dim hs As HighlightSet = GetOrCreateHighlightSet()
        If hs Is Nothing OrElse occ Is Nothing Then Return
        If occ.Definition Is Nothing OrElse occ.Definition.Document Is Nothing Then Return
        Dim key As String = occ.Definition.Document.FullDocumentName
        Dim quitar As Boolean = OccEstaResaltada(hs, occ)
        Dim doc As Document = _app.ActiveDocument
        If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then Return
        Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
        RecorrerToggleHighlight(asmDoc.ComponentDefinition.Occurrences, key, hs, quitar)
        _app.ActiveView.Update()
        HabilitarVentanaInventor()
    End Sub

    Private Sub RecorrerToggleHighlight(occs As ComponentOccurrences, key As String, hs As HighlightSet, quitar As Boolean)
        For Each occ As ComponentOccurrence In occs
            If occ.Suppressed Then Continue For
            If occ.Definition IsNot Nothing AndAlso occ.Definition.Document IsNot Nothing Then
                If occ.Definition.Document.FullDocumentName = key Then
                    If quitar Then
                        Try : hs.Remove(occ) : Catch : End Try
                    Else
                        Try : hs.AddItem(occ) : Catch : End Try
                    End If
                End If
            End If
            If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                RecorrerToggleHighlight(occ.SubOccurrences, key, hs, quitar)
            End If
        Next
    End Sub

    Private Sub SyncHighlightWithCheckbox(occ As ComponentOccurrence, checked As Boolean)
        Dim hs As HighlightSet = GetOrCreateHighlightSet()
        If hs Is Nothing Then Return
        Dim estaResaltada As Boolean = OccEstaResaltada(hs, occ)
        If checked AndAlso Not estaResaltada Then
            Try : hs.AddItem(occ) : Catch : End Try
        ElseIf Not checked AndAlso estaResaltada Then
            Try : hs.Remove(occ) : Catch : End Try
        End If
        _app.ActiveView.Update()
        HabilitarVentanaInventor()
    End Sub

    Private Function OccEstaResaltada(hs As HighlightSet, occ As ComponentOccurrence) As Boolean
        Try
            For Each item As Object In hs
                If item Is occ Then Return True
            Next
        Catch
        End Try
        Return False
    End Function

    Public Sub ClearHighlight()
        If _highlightSet IsNot Nothing Then
            Try
                _highlightSet.Clear()
                _app.ActiveView.Update()
                HabilitarVentanaInventor()
            Catch
            End Try
        End If
    End Sub

    ' =========================================================
    ' MANEJO DE CAMBIOS Y CIERRE
    ' =========================================================
    Private Sub ChkSoloConCambios_CheckedChanged(sender As Object, e As EventArgs)
        ActualizarPreview()
    End Sub

    Private Sub BtnRefrescar_Click(sender As Object, e As EventArgs)
        Try
            Dim doc As Document = _app.ActiveDocument
            If doc IsNot Nothing AndAlso doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                doc.SelectSet.Clear()
                ClearHighlight()
                _app.ActiveView.Update()
            End If
        Catch
        End Try
        CargarDatos()
    End Sub

    Private Sub BtnAplicar_Click(sender As Object, e As EventArgs)
        Dim cambios = _treeData.Where(Function(n) n.TieneCambio AndAlso Not n.IsBroken AndAlso Not n.IsBuloneria).ToList()
        If cambios.Count = 0 Then
            MessageBox.Show("No hay cambios para aplicar.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim previewMsg As String = String.Join(vbCrLf, cambios.Select(Function(n) n.FileNameShort & ": " & If(n.ComActual = "", "(vacío)", n.ComActual) & " → " & If(n.ComNuevo = "", "(vacío)", n.ComNuevo)).Take(20))
        If cambios.Count > 20 Then
            previewMsg &= vbCrLf & "... y " & (cambios.Count - 20) & " más"
        End If

        If MessageBox.Show("¿Aplicar estos cambios?" & vbCrLf & vbCrLf & previewMsg, "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Dim engine As New ComercialAssignmentEngine(_app)
            Dim result = engine.ExecuteWithChanges(cambios)
            If result.Success Then
                For Each nodo In cambios
                    nodo.ComActual = nodo.ComNuevo
                    nodo.TieneCambio = False
                Next

                For Each nodo In cambios
                    Dim treeNode As TreeNode = BuscarNodoPorDatos(tvArbol.Nodes, nodo)
                    If treeNode IsNot Nothing Then
                        RecalcularEstadoHaciaArriba(treeNode)
                    End If
                Next

                ClearHighlight()
                HighlightByComValue()
                ActualizarPreview()
                MessageBox.Show(String.Format("Completado.{0}Procesados: {1}{0}Fallidos: {2}", vbCrLf, result.ProcessedCount, result.FailedCount), "OK", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Error: " & result.ErrorMessage, "Fallo", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Function BuscarNodoPorDatos(nodes As TreeNodeCollection, datos As NodoComercial) As TreeNode
        For Each node As TreeNode In nodes
            If node.Tag IsNot Nothing AndAlso CType(node.Tag, NodoComercial) Is datos Then Return node
            If node.Nodes.Count > 0 Then
                Dim found As TreeNode = BuscarNodoPorDatos(node.Nodes, datos)
                If found IsNot Nothing Then Return found
            End If
        Next
        Return Nothing
    End Function

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs)
        Try
            Dim doc As Document = _app.ActiveDocument
            If doc IsNot Nothing AndAlso doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                doc.SelectSet.Clear()
                ClearHighlight()
                _app.ActiveView.Update()
            End If
        Catch
        End Try
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        Try
            Dim doc As Document = _app.ActiveDocument
            If doc IsNot Nothing AndAlso doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                doc.SelectSet.Clear()
                ClearHighlight()
                _app.ActiveView.Update()
            End If
        Catch
        End Try
        MyBase.OnFormClosing(e)
    End Sub
End Class

' =========================================================
' MODELO DE DATOS
' =========================================================
Public Class NodoComercial
    Public Property OccurrenceName As String
    Public Property FileName As String
    Public Property FileNameShort As String
    Public Property ComActual As String
    Public Property ComNuevo As String
    Public Property IsContentCenter As Boolean
    Public Property IsBroken As Boolean
    Public Property TieneCambio As Boolean
    Public Property IsBuloneria As Boolean
    Public Property Occurrence As ComponentOccurrence
    Public Property DefinitionType As DocumentTypeEnum
End Class

' =========================================================
' ENUMS Y CLASES DE APOYO PARA SELECCIÓN EN MODELO
' =========================================================
Public Enum ModoAccionModelo
    Alternar = 0
    Marcar = 1
    Desmarcar = 2
End Enum

Public Class ResultadoClickModelo
    Public Property Mensaje As String = ""
    Public Property CantidadAfectada As Integer = 0
    Public Property EsMarcado As Boolean = False
End Class

' =========================================================
' HELPER DE SELECCIÓN EN MODELO (BARRA FLOTANTE INTERACTIVA Y CONTINUA)
' =========================================================
Public Class SeleccionModeloHelper
    Inherits Form

    Private ReadOnly _app As Inventor.Application
    Private ReadOnly _parentDialog As ComercialAssignmentDialog
    Private ReadOnly _modoInicial As ModoAccionModelo
    Private _conteoModificados As Integer = 0

    ' Eventos nativos de Inventor
    Private WithEvents _interactionEvents As InteractionEvents
    Private WithEvents _selectEvents As SelectEvents

    ' Controles del Diálogo
    Private rbAlternar As RadioButton
    Private rbMarcar As RadioButton
    Private rbDesmarcar As RadioButton
    Private rbPieza As RadioButton
    Private rbEnsamblaje As RadioButton
    Private lblStatus As Label
    Private lblContador As Label
    Private btnFinalizar As Button

    Public Sub New(app As Inventor.Application, parentDialog As ComercialAssignmentDialog, Optional modoInicial As ModoAccionModelo = ModoAccionModelo.Alternar)
        _app = app
        _parentDialog = parentDialog
        _modoInicial = modoInicial

        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "🎯 Selección en Modelo 3D (Click continuo en piezas)"
        Me.Size = New System.Drawing.Size(530, 205)
        Me.FormBorderStyle = FormBorderStyle.FixedToolWindow
        Me.TopMost = True
        Me.StartPosition = FormStartPosition.Manual

        Dim screenArea As Rectangle = Screen.PrimaryScreen.WorkingArea
        Me.Location = New System.Drawing.Point((screenArea.Width - 530) \ 2, 40)
        Me.BackColor = System.Drawing.Color.FromArgb(248, 250, 252)
        Me.Font = New Font("Segoe UI", 9.0F)
        Me.ControlBox = True

        ' Panel 1: Modo / Acción (Alternar, Marcar, Desmarcar)
        Dim pnlModo As New Panel() With {
            .Location = New System.Drawing.Point(12, 10),
            .Size = New System.Drawing.Size(490, 34),
            .BackColor = System.Drawing.Color.FromArgb(241, 245, 249),
            .BorderStyle = BorderStyle.FixedSingle
        }
        Dim lblModo As New Label() With {
            .Text = "Acción:",
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(30, 41, 59),
            .Location = New System.Drawing.Point(8, 7),
            .AutoSize = True
        }
        pnlModo.Controls.Add(lblModo)

        rbAlternar = New RadioButton() With {
            .Text = "⇄ Alternar (Toggle)",
            .Location = New System.Drawing.Point(68, 5),
            .Size = New System.Drawing.Size(135, 22),
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(2, 132, 199),
            .Checked = (_modoInicial = ModoAccionModelo.Alternar),
            .Cursor = Cursors.Hand
        }
        pnlModo.Controls.Add(rbAlternar)

        rbMarcar = New RadioButton() With {
            .Text = "✓ Marcar (+)",
            .Location = New System.Drawing.Point(212, 5),
            .Size = New System.Drawing.Size(115, 22),
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(16, 149, 102),
            .Checked = (_modoInicial = ModoAccionModelo.Marcar),
            .Cursor = Cursors.Hand
        }
        pnlModo.Controls.Add(rbMarcar)

        rbDesmarcar = New RadioButton() With {
            .Text = "✕ Desmarcar (-)",
            .Location = New System.Drawing.Point(335, 5),
            .Size = New System.Drawing.Size(140, 22),
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(220, 38, 38),
            .Checked = (_modoInicial = ModoAccionModelo.Desmarcar),
            .Cursor = Cursors.Hand
        }
        pnlModo.Controls.Add(rbDesmarcar)
        Me.Controls.Add(pnlModo)

        ' Panel 2: Filtro de selección (Pieza vs Ensamblaje)
        Dim pnlFiltro As New Panel() With {
            .Location = New System.Drawing.Point(12, 48),
            .Size = New System.Drawing.Size(490, 28)
        }
        Dim lblFiltro As New Label() With {
            .Text = "Filtro:",
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(71, 85, 105),
            .Location = New System.Drawing.Point(8, 5),
            .AutoSize = True
        }
        pnlFiltro.Controls.Add(lblFiltro)

        rbPieza = New RadioButton() With {
            .Text = "Pieza individual (.ipt)",
            .Location = New System.Drawing.Point(68, 3),
            .Size = New System.Drawing.Size(160, 20),
            .Font = New Font("Segoe UI", 8.25F),
            .Checked = True,
            .Cursor = Cursors.Hand
        }
        AddHandler rbPieza.CheckedChanged, AddressOf OnFiltroChanged
        pnlFiltro.Controls.Add(rbPieza)

        rbEnsamblaje = New RadioButton() With {
            .Text = "Sub-ensamblaje (.iam)",
            .Location = New System.Drawing.Point(235, 3),
            .Size = New System.Drawing.Size(180, 20),
            .Font = New Font("Segoe UI", 8.25F),
            .Cursor = Cursors.Hand
        }
        AddHandler rbEnsamblaje.CheckedChanged, AddressOf OnFiltroChanged
        pnlFiltro.Controls.Add(rbEnsamblaje)
        Me.Controls.Add(pnlFiltro)

        ' Panel 3: Feedback en vivo y Contador
        lblStatus = New Label() With {
            .Text = "Haga click en las piezas en el visor 3D para aplicar la acción elegida. Presione ESC para finalizar.",
            .Location = New System.Drawing.Point(15, 82),
            .Size = New System.Drawing.Size(485, 30),
            .Font = New Font("Segoe UI", 8.25F, FontStyle.Italic),
            .ForeColor = System.Drawing.Color.FromArgb(51, 65, 85)
        }
        Me.Controls.Add(lblStatus)

        lblContador = New Label() With {
            .Text = "Piezas modificadas en esta sesión: 0",
            .Location = New System.Drawing.Point(15, 122),
            .Size = New System.Drawing.Size(265, 24),
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(15, 23, 42)
        }
        Me.Controls.Add(lblContador)

        btnFinalizar = New Button() With {
            .Text = "✓ Finalizar Selección (ESC)",
            .Location = New System.Drawing.Point(285, 118),
            .Size = New System.Drawing.Size(217, 34),
            .BackColor = System.Drawing.Color.FromArgb(2, 132, 199),
            .ForeColor = System.Drawing.Color.White,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        btnFinalizar.FlatAppearance.BorderSize = 0
        AddHandler btnFinalizar.Click, Sub() CerrarTodo()
        Me.Controls.Add(btnFinalizar)

        Me.AcceptButton = btnFinalizar
        Me.CancelButton = btnFinalizar
    End Sub

    ''' <summary>
    ''' Arranca el entorno interactivo en Inventor habilitando clicks continuos en el modelo 3D.
    ''' </summary>
    Public Sub IniciarSeleccion()
        Try
            _interactionEvents = _app.CommandManager.CreateInteractionEvents()
            _interactionEvents.StatusBarText = "Haga click en las piezas en el modelo 3D para marcar o desmarcar. Presione ESC para finalizar."

            _selectEvents = _interactionEvents.SelectEvents
            _selectEvents.SingleSelectEnabled = False ' Mantiene la selección activa para hacer múltiples clicks sucesivos

            ActualizarFiltroDeSeleccion()

            _interactionEvents.Start()
            Me.Show()
            _parentDialog.HabilitarVentanaInventor()
        Catch ex As Exception
            MessageBox.Show("Error al iniciar los eventos de selección: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CerrarTodo()
        End Try
    End Sub

    Private Sub OnFiltroChanged(sender As Object, e As EventArgs)
        If _selectEvents Is Nothing Then Return
        ActualizarFiltroDeSeleccion()
    End Sub

    Private Sub ActualizarFiltroDeSeleccion()
        If _selectEvents Is Nothing Then Return
        _selectEvents.ClearSelectionFilter()

        If rbPieza.Checked Then
            _selectEvents.AddSelectionFilter(SelectionFilterEnum.kPartFeatureFilter)
            _selectEvents.AddSelectionFilter(SelectionFilterEnum.kAssemblyLeafOccurrenceFilter)
        Else
            _selectEvents.AddSelectionFilter(SelectionFilterEnum.kAssemblyOccurrenceFilter)
        End If
    End Sub

    Private Function ObtenerModoActual() As ModoAccionModelo
        If rbDesmarcar.Checked Then Return ModoAccionModelo.Desmarcar
        If rbMarcar.Checked Then Return ModoAccionModelo.Marcar
        Return ModoAccionModelo.Alternar
    End Function

    ''' <summary>
    ''' Evento nativo de click en el 3D de Inventor.
    ''' Procesa inmediatamente el componente sin cerrar el diálogo ni bloquear la vista.
    ''' </summary>
    Private Sub _selectEvents_OnSelect(ByVal JustSelectedEntities As ObjectsEnumerator, ByVal SelectionDevice As SelectionDeviceEnum, ByVal ModelPosition As Inventor.Point, ByVal ViewPosition As Point2d, ByVal View As Inventor.View) Handles _selectEvents.OnSelect
        Try
            If JustSelectedEntities IsNot Nothing AndAlso JustSelectedEntities.Count > 0 Then
                Dim obj As Object = JustSelectedEntities.Item(1)
                Dim occ As ComponentOccurrence = Nothing

                If TypeOf obj Is ComponentOccurrence Then
                    occ = CType(obj, ComponentOccurrence)
                Else
                    Try
                        occ = obj.ContainingOccurrence
                    Catch
                    End Try
                End If

                If occ IsNot Nothing Then
                    Dim modo As ModoAccionModelo = ObtenerModoActual()
                    Dim recursivo As Boolean = rbEnsamblaje.Checked
                    Dim res = _parentDialog.ProcesarClickModeloEnVivo(occ, modo, recursivo)

                    lblStatus.Text = res.Mensaje
                    If res.CantidadAfectada > 0 Then
                        _conteoModificados += res.CantidadAfectada
                        lblContador.Text = String.Format("Piezas modificadas en esta sesión: {0}", _conteoModificados)
                    End If
                End If
            End If
        Catch ex As Exception
        Finally
            Try
                Dim doc As Document = _app.ActiveDocument
                If doc IsNot Nothing Then doc.SelectSet.Clear()
            Catch
            End Try
        End Try
    End Sub

    Private Sub _interactionEvents_OnTerminate() Handles _interactionEvents.OnTerminate
        CerrarTodo()
    End Sub

    Private Sub CerrarTodo()
        Try
            If _interactionEvents IsNot Nothing Then
                _interactionEvents.Stop()
            End If
        Catch
        Finally
            _selectEvents = Nothing
            _interactionEvents = Nothing
            Me.Hide()
            _parentDialog.OnSeleccionModeloFinalizada()
            Me.Close()
        End Try
    End Sub
End Class