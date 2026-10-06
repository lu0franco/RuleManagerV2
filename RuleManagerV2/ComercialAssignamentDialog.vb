Imports System.Windows.Forms
Imports System.Drawing
Imports System.Linq
Imports Inventor

''' <summary>
''' Diálogo compacto de asignación COM con indicador recursivo vía StateImageList.
''' Jerarquía visual: [StateImage toggle] [Image tipo] [Texto]
''' </summary>
Public Class ComercialAssignmentDialog
    Inherits Form

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
    Private btnHighlightCom As Button
    Private chkSoloConCambios As CheckBox

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
        Me.Text = "Asignación COM"
        Me.Size = New System.Drawing.Size(680, 700)
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New System.Drawing.Point(50, 50)
        Me.MinimumSize = New System.Drawing.Size(500, 400)
        Me.FormBorderStyle = FormBorderStyle.Sizable

        ' --- Panel izquierdo: TreeView ---
        Dim panelArbol As New Panel()
        panelArbol.Dock = DockStyle.Fill

        Dim lblArbol As New Label()
        lblArbol.Text = "[☑]=Tiene COM  [▣]=Parcial  [□]=Sin COM  Click=Toggle  Doble click=zoom"
        lblArbol.Dock = DockStyle.Top
        lblArbol.Height = 24
        lblArbol.Padding = New Padding(5, 3, 5, 3)
        lblArbol.Font = New Font("Segoe UI", 8)
        lblArbol.BackColor = System.Drawing.Color.FromArgb(240, 240, 240)

        ' Crear ImageLists
        _imageList = CrearImageList()
        _stateImageList = CrearStateImageList()

        tvArbol = New TreeView()
        tvArbol.Dock = DockStyle.Fill
        tvArbol.CheckBoxes = False
        tvArbol.ImageList = _imageList
        tvArbol.StateImageList = _stateImageList
        tvArbol.FullRowSelect = True

        AddHandler tvArbol.NodeMouseClick, AddressOf TvArbol_NodeMouseClick
        AddHandler tvArbol.NodeMouseDoubleClick, AddressOf TvArbol_NodeMouseDoubleClick
        AddHandler tvArbol.MouseClick, AddressOf TvArbol_MouseClick

        panelArbol.Controls.Add(tvArbol)
        panelArbol.Controls.Add(lblArbol)

        ' --- Panel derecho: Botones ---
        Dim panelBotones As New Panel()
        panelBotones.Dock = DockStyle.Right
        panelBotones.Width = 150
        panelBotones.BackColor = System.Drawing.Color.FromArgb(248, 248, 248)
        panelBotones.Padding = New Padding(8, 10, 8, 10)

        Dim y As Integer = 10

        Dim lblTitulo As New Label()
        lblTitulo.Text = "Acciones"
        lblTitulo.Location = New System.Drawing.Point(8, y)
        lblTitulo.Size = New System.Drawing.Size(134, 20)
        lblTitulo.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        panelBotones.Controls.Add(lblTitulo)
        y += 28

        btnSeleccionarModelo = CrearBotonLateral("Seleccionar del modelo", y, System.Drawing.Color.FromArgb(0, 120, 80))
        AddHandler btnSeleccionarModelo.Click, AddressOf BtnSeleccionarModelo_Click
        panelBotones.Controls.Add(btnSeleccionarModelo)
        y += 42

        btnHighlightCom = CrearBotonLateral("Resaltar COM", y, System.Drawing.Color.FromArgb(200, 0, 0))
        AddHandler btnHighlightCom.Click, AddressOf BtnHighlightCom_Click
        panelBotones.Controls.Add(btnHighlightCom)
        y += 50

        y += 10
        Dim sep1 As New Label()
        sep1.Location = New System.Drawing.Point(8, y)
        sep1.Size = New System.Drawing.Size(134, 2)
        sep1.BorderStyle = BorderStyle.Fixed3D
        panelBotones.Controls.Add(sep1)
        y += 15

        chkSoloConCambios = New CheckBox()
        chkSoloConCambios.Text = "Solo cambios"
        chkSoloConCambios.Location = New System.Drawing.Point(8, y)
        chkSoloConCambios.Size = New System.Drawing.Size(134, 20)
        chkSoloConCambios.Font = New Font("Segoe UI", 8)
        chkSoloConCambios.AutoSize = False
        AddHandler chkSoloConCambios.CheckedChanged, AddressOf ChkSoloConCambios_CheckedChanged
        panelBotones.Controls.Add(chkSoloConCambios)
        y += 30

        btnRefrescar = CrearBotonLateral("Refrescar", y, System.Drawing.Color.FromArgb(100, 100, 100))
        AddHandler btnRefrescar.Click, AddressOf BtnRefrescar_Click
        panelBotones.Controls.Add(btnRefrescar)
        y += 50

        y += 10
        Dim sep2 As New Label()
        sep2.Location = New System.Drawing.Point(8, y)
        sep2.Size = New System.Drawing.Size(134, 2)
        sep2.BorderStyle = BorderStyle.Fixed3D
        panelBotones.Controls.Add(sep2)
        y += 15

        btnAplicar = CrearBotonLateral("Aplicar cambios", y, System.Drawing.Color.FromArgb(0, 120, 212))
        AddHandler btnAplicar.Click, AddressOf BtnAplicar_Click
        panelBotones.Controls.Add(btnAplicar)
        y += 42

        btnCancelar = CrearBotonLateral("Cancelar", y, System.Drawing.Color.FromArgb(120, 120, 120))
        AddHandler btnCancelar.Click, AddressOf BtnCancelar_Click
        panelBotones.Controls.Add(btnCancelar)
        y += 42

        y += 15
        Dim sep3 As New Label()
        sep3.Location = New System.Drawing.Point(8, y)
        sep3.Size = New System.Drawing.Size(134, 2)
        sep3.BorderStyle = BorderStyle.Fixed3D
        panelBotones.Controls.Add(sep3)
        y += 12

        lblResumen = New Label()
        lblResumen.Text = "Listo"
        lblResumen.Location = New System.Drawing.Point(8, y)
        lblResumen.Size = New System.Drawing.Size(134, 80)
        lblResumen.Font = New Font("Segoe UI", 8)
        lblResumen.ForeColor = System.Drawing.Color.DimGray
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
        btn.Size = New System.Drawing.Size(134, 36)
        btn.BackColor = colorFondo
        btn.ForeColor = System.Drawing.Color.White
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 0
        btn.Font = New Font("Segoe UI", 8.5, FontStyle.Bold)
        btn.TextAlign = ContentAlignment.MiddleCenter
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
            Dim cam As Camera = view.Camera
            Dim box As Box = occ.RangeBox

            Dim center As Inventor.Point = _app.TransientGeometry.CreateSystem.Drawing.Point(
                (box.MinSystem.Drawing.Point.X + box.MaxSystem.Drawing.Point.X) / 2,
                (box.MinSystem.Drawing.Point.Y + box.MaxSystem.Drawing.Point.Y) / 2,
                (box.MinSystem.Drawing.Point.Z + box.MaxSystem.Drawing.Point.Z) / 2)

            Dim sizeX As Double = Math.Abs(box.MaxSystem.Drawing.Point.X - box.MinSystem.Drawing.Point.X)
            Dim sizeY As Double = Math.Abs(box.MaxSystem.Drawing.Point.Y - box.MinSystem.Drawing.Point.Y)
            Dim sizeZ As Double = Math.Abs(box.MaxSystem.Drawing.Point.Z - box.MinSystem.Drawing.Point.Z)
            Dim maxSize As Double = Math.Max(sizeX, Math.Max(sizeY, sizeZ))
            Dim distance As Double = Math.Max(maxSize * 2.5, 10)

            cam.TargetSystem.Drawing.Point = center
            cam.EyeSystem.Drawing.Point = _app.TransientGeometry.CreateSystem.Drawing.Point(
                center.X + distance,
                center.Y + distance,
                center.Z + distance)
            cam.UpVector = _app.TransientGeometry.CreateUnitVector(0, 0, 1)
            cam.Apply()

            doc.SelectSet.Clear()
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
    ' SELECCION MODELO CON COMMANDPICKER NATIVO (INTERACTIONEVENTS)
    ' =========================================================
    Private Sub BtnSeleccionarModelo_Click(sender As Object, e As EventArgs)
        Try
            Dim doc As Document = _app.ActiveDocument
            If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("No hay un ensamblaje activo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            doc.SelectSet.Clear()
            Me.Hide() ' Ocultamos este diálogo para interactuar con Inventor

            ' Inicializamos el asistente nativo de picking
            Dim helper As New SeleccionModeloHelper(_app, AddressOf OnSeleccionModeloCompletada)
            helper.IniciarSeleccion()

        Catch ex As Exception
            Me.Show()
            Me.Activate()
            MessageBox.Show("Error al iniciar selección nativa: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' FLUJO AUTOMATIZADO: Al aceptar la selección, marca, aplica, refresca y resalta automáticamente.
    ''' </summary>
    Private Sub OnSeleccionModeloCompletada(occ As ComponentOccurrence, cancelado As Boolean)
        Me.Show()
        Me.Activate()
        Me.WindowState = FormWindowState.Normal

        If cancelado OrElse occ Is Nothing Then Return

        Try
            Dim doc As Document = _app.ActiveDocument
            doc.SelectSet.Clear()

            Dim esIAM As Boolean = (occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject)
            Dim modoRecursivo As Boolean = False

            If esIAM Then
                Dim respuesta As DialogResult = MessageBox.Show(
                    "Seleccionó el sub-ensamblaje: " & occ.Name & vbCrLf & vbCrLf &
                    "¿Desea marcar COM recursivamente a todos sus componentes internos?" & vbCrLf &
                    "• Sí = Todo el contenido interno heredará la marca." & vbCrLf &
                    "• No = Solo el contenedor .iam se marcará.",
                    "Modo de selección .IAM",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question)

                If respuesta = DialogResult.Cancel Then Return
                modoRecursivo = (respuesta = DialogResult.Yes)
            End If

            ' 1. Buscar y simular la marca en nuestra lista en memoria
            Dim nodoEncontrado As TreeNode = BuscarNodoPorOccurrence(tvArbol.Nodes, occ)
            If nodoEncontrado IsNot Nothing Then
                If modoRecursivo Then
                    MarcarComRecursivo(nodoEncontrado, True)
                Else
                    ProcesarNodoComRecursivo(nodoEncontrado, False)
                End If
            Else
                ProcesarOcurrenciaComRecursivo(occ, modoRecursivo)
            End If

            ' 2. AUTOMATIZACIÓN TRIPLE: Aplicar cambios inmediatamente sin preguntar
            Dim cambios = _treeData.Where(Function(n) n.TieneCambio AndAlso Not n.IsBroken AndAlso Not n.IsBuloneria).ToList()
            If cambios.Count > 0 Then
                Dim engine As New ComercialAssignmentEngine(_app)
                Dim result = engine.ExecuteWithChanges(cambios)

                If result.Success Then
                    ' Sincronizar estados
                    For Each nodo In cambios
                        nodo.ComActual = nodo.ComNuevo
                        nodo.TieneCambio = False
                    Next
                End If
            End If

            ' 3. Refrescar la UI leyendo el modelo real
            CargarDatos()

            ' 4. Volver a encender el Resaltado de todo lo que tenga COM activo
            ClearHighlight()
            HighlightByComValue()

        Catch ex As Exception
            MessageBox.Show("Error en procesamiento automático: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function BuscarNodoPorOccurrence(nodes As TreeNodeCollection, occ As ComponentOccurrence) As TreeNode
        For Each node As TreeNode In nodes
            If node.Tag IsNot Nothing Then
                Dim datos As NodoComercial = CType(node.Tag, NodoComercial)
                If datos.Occurrence Is occ Then Return node
            End If
            If node.Nodes.Count > 0 Then
                Dim found As TreeNode = BuscarNodoPorOccurrence(node.Nodes, occ)
                If found IsNot Nothing Then Return found
            End If
        Next
        Return Nothing
    End Function

    Private Sub ProcesarNodoComRecursivo(nodo As TreeNode, recursivo As Boolean)
        If nodo Is Nothing Then Return
        If nodo.Tag IsNot Nothing Then
            Dim datos As NodoComercial = CType(nodo.Tag, NodoComercial)
            If Not datos.IsContentCenter AndAlso Not datos.IsBuloneria AndAlso Not datos.IsBroken Then
                If nodo.StateImageIndex <> 1 Then
                    nodo.StateImageIndex = 1
                    datos.ComNuevo = "✓"
                    datos.TieneCambio = (datos.ComActual <> datos.ComNuevo)
                    If datos.Occurrence IsNot Nothing Then
                        SyncHighlightWithCheckbox(datos.Occurrence, True)
                    End If
                End If
            End If
        End If

        If recursivo Then
            For Each hijo As TreeNode In nodo.Nodes
                ProcesarNodoComRecursivo(hijo, True)
            Next
        End If
        RecalcularEstadoHaciaArriba(nodo)
    End Sub

    Private Sub ProcesarOcurrenciaComRecursivo(occ As ComponentOccurrence, recursivo As Boolean)
        If occ Is Nothing Then Return
        Try
            If occ.Definition IsNot Nothing AndAlso occ.Definition.Document IsNot Nothing Then
                Dim esBuloneriaValor As Boolean = EsBuloneria(occ.Definition.Document)
                Dim esCC As Boolean = (occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject AndAlso occ.Definition.IsContentMember)
                If Not esBuloneriaValor AndAlso Not esCC Then
                    Dim fileName As String = occ.Definition.Document.FullDocumentName
                    Dim datos As NodoComercial = _treeData.FirstOrDefault(Function(n) n.FileName.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    If datos IsNot Nothing Then
                        datos.ComNuevo = "✓"
                        datos.TieneCambio = (datos.ComActual <> datos.ComNuevo)
                    End If
                    SyncHighlightWithCheckbox(occ, True)
                End If
            End If

            If recursivo AndAlso occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                For Each subOcc As ComponentOccurrence In occ.SubOccurrences
                    ProcesarOcurrenciaComRecursivo(subOcc, True)
                Next
            End If
        Catch
        End Try
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
' HELPER DE SELECCIÓN EN MODELO (DIÁLOGO FLOTANTE CON COMMANDPICKER)
' =========================================================
Public Class SeleccionModeloHelper
    Inherits Form

    Private _app As Inventor.Application
    Private _callback As Action(Of ComponentOccurrence, Boolean)

    ' Eventos nativos de Inventor
    Private WithEvents _interactionEvents As InteractionEvents
    Private WithEvents _selectEvents As SelectEvents

    ' Controles del Diálogo
    Private rbPieza As RadioButton
    Private rbEnsamblaje As RadioButton
    Private btnAceptar As Button
    Private btnCancelar As Button
    Private lblInfo As Label

    Private _occSeleccionada As ComponentOccurrence = Nothing

    Public Sub New(app As Inventor.Application, callback As Action(Of ComponentOccurrence, Boolean))
        _app = app
        _callback = callback

        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Selección desde el Modelo"
        Me.Size = New System.Drawing.Size(320, 190)
        Me.FormBorderStyle = FormBorderStyle.FixedToolWindow
        Me.TopMost = True
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.ControlBox = False ' Evita que cierren con la X sin cancelar eventos

        lblInfo = New Label()
        lblInfo.Text = "1. Elija qué tipo de componente desea buscar:"
        lblInfo.Location = New System.Drawing.Point(12, 10)
        lblInfo.Size = New System.Drawing.Size(280, 20)
        lblInfo.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        Me.Controls.Add(lblInfo)

        ' Modo Pieza (.ipt)
        rbPieza = New RadioButton()
        rbPieza.Text = "Pieza individual (.ipt)"
        rbPieza.Location = New System.Drawing.Point(25, 35)
        rbPieza.Size = New System.Drawing.Size(200, 20)
        rbPieza.Checked = True
        AddHandler rbPieza.CheckedChanged, AddressOf OnFiltroChanged
        Me.Controls.Add(rbPieza)

        ' Modo Ensamblaje (.iam)
        rbEnsamblaje = New RadioButton()
        rbEnsamblaje.Text = "Sub-ensamblaje contenedor (.iam)"
        rbEnsamblaje.Location = New System.Drawing.Point(25, 58)
        rbEnsamblaje.Size = New System.Drawing.Size(250, 20)
        AddHandler rbEnsamblaje.CheckedChanged, AddressOf OnFiltroChanged
        Me.Controls.Add(rbEnsamblaje)

        ' Botón Aceptar
        btnAceptar = New Button()
        btnAceptar.Text = "Aceptar"
        btnAceptar.Location = New System.Drawing.Point(40, 105)
        btnAceptar.Size = New System.Drawing.Size(110, 32)
        btnAceptar.BackColor = System.Drawing.Color.FromArgb(0, 120, 80)
        btnAceptar.ForeColor = System.Drawing.Color.White
        btnAceptar.FlatStyle = FlatStyle.Flat
        btnAceptar.FlatAppearance.BorderSize = 0
        btnAceptar.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        AddHandler btnAceptar.Click, AddressOf BtnAceptar_Click
        Me.Controls.Add(btnAceptar)

        ' Botón Cancelar
        btnCancelar = New Button()
        btnCancelar.Text = "Cancelar"
        btnCancelar.Location = New System.Drawing.Point(160, 105)
        btnCancelar.Size = New System.Drawing.Size(110, 32)
        btnCancelar.BackColor = System.Drawing.Color.FromArgb(120, 120, 120)
        btnCancelar.ForeColor = System.Drawing.Color.White
        btnCancelar.FlatStyle = FlatStyle.Flat
        btnCancelar.FlatAppearance.BorderSize = 0
        btnCancelar.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        AddHandler btnCancelar.Click, AddressOf BtnCancelar_Click
        Me.Controls.Add(btnCancelar)
    End Sub

    ''' <summary>
    ''' Arranca el entorno interactivo en Inventor al mostrar el formulario.
    ''' </summary>
    Public Sub IniciarSeleccion()
        Try
            _interactionEvents = _app.CommandManager.CreateInteractionEvents()
            _interactionEvents.StatusBarText = "Haga click en el componente del modelo 3D según el filtro seleccionado."

            _selectEvents = _interactionEvents.SelectEvents
            _selectEvents.SingleSelectEnabled = True

            ' Configurar filtro inicial
            ActualizarFiltroDeSeleccion()

            _interactionEvents.Start()
            Me.Show()
        Catch ex As Exception
            MessageBox.Show("Error al iniciar los eventos de selección: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CerrarTodo(Nothing, True)
        End Try
    End Sub

    ''' <summary>
    ''' Cambia dinámicamente el comportamiento del cursor de Inventor sin romper el comando
    ''' </summary>
    Private Sub OnFiltroChanged(sender As Object, e As EventArgs)
        If _selectEvents Is Nothing Then Return
        ActualizarFiltroDeSeleccion()
    End Sub

    Private Sub ActualizarFiltroDeSeleccion()
        _selectEvents.ClearSelectionFilter()

        If rbPieza.Checked Then
            ' kPartAppearanceFilter / LeafObect fuerza a buscar la pieza de nivel más bajo (.ipt)
            _selectEvents.AddSelectionFilter(SelectionFilterEnum.kPartFeatureFilter)
            _selectEvents.AddSelectionFilter(SelectionFilterEnum.kAssemblyLeafOccurrenceFilter)
        Else
            ' kAssemblyOccurrenceFilter permite agarrar sub-ensamblajes de niveles superiores (.iam)
            _selectEvents.AddSelectionFilter(SelectionFilterEnum.kAssemblyOccurrenceFilter)
        End If
    End Sub

    ''' <summary>
    ''' Evento nativo de click en el 3D de Inventor
    ''' </summary>
    Private Sub _selectEvents_OnSelect(ByVal JustSelectedEntities As ObjectsEnumerator, ByVal SelectionDevice As SelectionDeviceEnum, ByVal ModelPosition As Inventor.Point, ByVal ViewPosition As Point2d, ByVal View As Inventor.View) Handles _selectEvents.OnSelect
        Try
            If JustSelectedEntities IsNot Nothing AndAlso JustSelectedEntities.Count > 0 Then
                Dim obj As Object = JustSelectedEntities.Item(1)

                ' Si es una cara/arista de pieza, escalamos hasta su ocurrencia
                If TypeOf obj Is ComponentOccurrence Then
                    _occSeleccionada = CType(obj, ComponentOccurrence)
                Else
                    ' Intentar obtener el objeto contenedor mediante la propiedad de la API si aplica
                    Try
                        _occSeleccionada = obj.ContainingOccurrence
                    Catch
                        _occSeleccionada = Nothing
                    End Try
                End If
            End If
        Catch
            _occSeleccionada = Nothing
        End Try
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs)
        If _occSeleccionada Is Nothing Then
            MessageBox.Show("Aún no ha hecho click en ningún componente del modelo válido.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        CerrarTodo(_occSeleccionada, False)
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs)
        CerrarTodo(Nothing, True)
    End Sub

    Private Sub _interactionEvents_OnTerminate() Handles _interactionEvents.OnTerminate
        ' Si el usuario presiona ESC en Inventor, actúa como Cancelar
        btnCancelar.PerformClick()
    End Sub

    Private Sub CerrarTodo(occ As ComponentOccurrence, cancelado As Boolean)
        Try
            If _interactionEvents IsNot Nothing Then
                _interactionEvents.Stop()
            End If
        Catch
        Finally
            _selectEvents = Nothing
            _interactionEvents = Nothing
            Me.Hide()
            _callback(occ, cancelado)
            Me.Close()
        End Try
    End Sub
End Class