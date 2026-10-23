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
Imports Size = System.Drawing.Size
Imports WinTextBox = System.Windows.Forms.TextBox
Imports TextBox = System.Windows.Forms.TextBox

''' <summary>
''' Representa un componente para el análisis y asignación de ensamble padre y numeración secuencial.
''' Diseñado con paridad total con la Lista de Materiales (BOM) Estructurada de Autodesk Inventor.
''' </summary>
Public Class ElementoPadreItem
    Public Property Occurrence As ComponentOccurrence
    Public Property Document As Document
    Public Property DocumentPath As String = ""
    Public Property FileName As String = ""
    Public Property Extension As String = ""
    Public Property ItemNumber As String = ""            ' Posición en BOM (1, 2, 3...)
    Public Property Quantity As Integer = 1             ' CTDAD en BOM
    Public Property BomStructure As String = "Normal"   ' Normal, Comprado, Inseparable, Fantasma
    Public Property IsAssembly As Boolean = False
    Public Property IsCommercial As Boolean = False
    Public Property IsContentCenter As Boolean = False
    Public Property Categoria As String = "CYP"         ' CYP, MEC, ARM, MC
    Public Property CurrentPadre As String = ""
    Public Property SuggestedPadre As String = ""
    Public Property CurrentPartNumber As String = ""
    Public Property CurrentStockNumber As String = ""   ' Nº de almacenamiento en Inventor
    Public Property ProposedPosition As String = ""     ' 01, 02, 03...
    Public Property ProposedCode As String = ""         ' Código resultante (e.g. OT-26282.04.09.01)
    Public Property ManualOverride As Boolean = False
    Public Property Level As Integer = 0
    Public Property Included As Boolean = True

    Public ReadOnly Property ColumnaTipo As String
        Get
            If IsCommercial Then
                Return "Comercial (COM)"
            ElseIf IsContentCenter Then
                Return "Librería (CC)"
            ElseIf IsAssembly Then
                Return "Ensamble"
            Else
                Return "Fabricación (" & Categoria & ")"
            End If
        End Get
    End Property

    Public ReadOnly Property CodigoComercial As String
        Get
            If Not String.IsNullOrEmpty(CurrentStockNumber) Then Return CurrentStockNumber
            If Not String.IsNullOrEmpty(CurrentPartNumber) Then Return CurrentPartNumber
            Return IOPath.GetFileNameWithoutExtension(FileName)
        End Get
    End Property

    Public ReadOnly Property IconoPrefix As String
        Get
            If IsCommercial Then
                Return "📦 "
            ElseIf IsContentCenter Then
                Return "🔩 "
            ElseIf IsAssembly Then
                Return "📁 "
            Else
                Return "📄 "
            End If
        End Get
    End Property

    Public ReadOnly Property RamaFormateada As String
        Get
            Dim indent As String = ""
            If Level > 1 Then
                indent = New String(" "c, (Level - 1) * 3) & "└── "
            End If
            Return indent & IconoPrefix & FileName
        End Get
    End Property

    Public ReadOnly Property TieneCambio As Boolean
        Get
            If Not Included OrElse String.IsNullOrEmpty(ProposedPosition) OrElse ProposedPosition = "-" Then Return False
            Return Not String.Equals(CurrentStockNumber, ProposedCode, StringComparison.OrdinalIgnoreCase) OrElse
                   Not String.Equals(CurrentPartNumber, ProposedCode, StringComparison.OrdinalIgnoreCase) OrElse
                   Not String.Equals(CurrentPadre, SuggestedPadre, StringComparison.OrdinalIgnoreCase)
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return If(Not String.IsNullOrEmpty(ProposedCode), ProposedCode, FileName) & " (Padre: " & SuggestedPadre & ")"
    End Function
End Class

''' <summary>
''' Diálogo profesional e intuitivo para numerar piezas en función del ensamble padre (ASIGNAR PADRES).
''' Diseñado con el mismo formato y filosofía visual de 'Asignar Propiedades' (SobremedidaLargoDialog):
''' - Agrupamiento jerárquico de piezas bajo su ensamble/subensamblaje en un árbol con ramas y carpetas (TreeView)
''' - Visualización sincronizada de árbol y grilla (filtrado por rama o navegación fluida)
''' - Totalmente NO MODAL (Modeless): permite rotar, orbitar y navegar libremente la ventana 3D de Inventor
''' - Sin dependencias externas incompatibles (stdole / IPictureDisp eliminadas, usa PictureDispConverter)
''' - Paridad 100% con la Lista de Materiales (BOM) estructurada de Inventor con fallback recursivo
''' - Resaltado 3D interactivo en color celeste vibrante con centrado de cámara
''' - Abrir elemento en Inventor o explorador con un solo clic
''' </summary>
Public Class AsignarPadresDialog
    Inherits Form

    <DllImport("user32.dll")>
    Private Shared Function EnableWindow(hWnd As IntPtr, bEnable As Boolean) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function

    Private ReadOnly _invApp As Inventor.Application
    Private ReadOnly _asmDoc As AssemblyDocument
    Private ReadOnly _items As New List(Of ElementoPadreItem)()
    Private _highlightSet As HighlightSet = Nothing
    Private _isSyncingSelection As Boolean = False
    Private _filtroRamaSeleccionada As TreeNode = Nothing

    ' Contenedores de diseño
    Private pnlTop As Panel
    Private pnlBottom As Panel
    Private pnlRight As Panel
    Private splitCenter As SplitContainer
    Private treeEstructura As TreeView
    Private _grid As DataGridView

    ' Controles de configuración y filtrado en panel superior
    Private txtBuscar As WinTextBox
    Private cmbEsquema As ComboBox
    Private cmbDigitos As ComboBox
    Private cmbSeparador As ComboBox
    Private chkOmitirComerciales As CheckBox
    Private chkOmitirContentCenter As CheckBox
    Private chkActualizarStockNumber As CheckBox
    Private chkActualizarPartNumber As CheckBox
    Private chkActualizarPadre As CheckBox
    Private chkActualizarPos As CheckBox
    Private chkTopMost As CheckBox
    Private chkFiltrarPorArbol As CheckBox
    Private btnAutoNumerar As Button
    Private btnMarcarTodos As Button
    Private btnDesmarcarTodos As Button

    ' Barra de herramientas del árbol
    Private btnTreeExpandAll As Button
    Private btnTreeCollapseAll As Button
    Private btnTreeReset As Button

    ' Controles del panel lateral derecho (Detalle / Inspección)
    Private picPreview As PictureBox
    Private lblDetalleNombre As Label
    Private lblDetalleTipo As Label
    Private lblDetallePadre As Label
    Private lblDetalleStockActual As Label
    Private lblDetallePartNumber As Label
    Private lblDetallePos As Label
    Private lblDetallePropuesto As Label
    Private btnZoom As Button
    Private btnAbrir As Button
    Private btnAceptarLateral As Button
    Private btnCancelarLateral As Button

    ' Controles del panel inferior
    Private lblResumen As Label
    Private btnAplicar As Button
    Private btnCancelar As Button

    Public ReadOnly Property Items As List(Of ElementoPadreItem)
        Get
            Return _items
        End Get
    End Property

    Public Sub New(invApp As Inventor.Application, asmDoc As AssemblyDocument)
        _invApp = invApp
        _asmDoc = asmDoc

        InitializeComponentsClean()
        CargarJerarquiaDesdeBOM()
        AutoDetectarSeparadorYPrefijo()
        AutoNumerarElementos()
        ConstruirArbolJerarquico()
        RefrescarGrilla()
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            Try
                LimpiarResaltado()
            Catch
            End Try
        End If
        MyBase.Dispose(disposing)
    End Sub

    ''' <summary>
    ''' Asegura que la ventana principal de Autodesk Inventor y su visor 3D permanezcan habilitados
    ''' para permitir rotar, orbitar y navegar el ensamble sin ninguna interferencia.
    ''' </summary>
    Public Sub HabilitarVentanaInventor()
        Try
            If _invApp IsNot Nothing Then
                Dim hwnd As New IntPtr(_invApp.MainFrameHWND)
                If hwnd <> IntPtr.Zero Then
                    EnableWindow(hwnd, True)
                End If
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Inicializa los componentes de la interfaz de usuario con diseño moderno y soporte
    ''' split-view entre el árbol de carpetas/subensamblajes y la grilla de propiedades.
    ''' </summary>
    Private Sub InitializeComponentsClean()
        Me.SuspendLayout()

        Me.Text = "ASIGNAR PADRES - Estructura Jerárquica y Numeración Sistemática"
        Me.Size = New Size(1220, 740)
        Me.MinimumSize = New Size(1000, 580)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(248, 250, 252)
        Me.ForeColor = Color.FromArgb(15, 23, 42)
        Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
        Me.TopMost = True

        ' Configuración de eventos para mantener Inventor 100% interactivo
        AddHandler Me.Shown, Sub() HabilitarVentanaInventor()
        AddHandler Me.Activated, Sub() HabilitarVentanaInventor()
        AddHandler Me.Deactivate, Sub() HabilitarVentanaInventor()

        ' =====================================================================
        ' 1. PANEL SUPERIOR: Título, metadatos del documento y opciones de numeración
        ' =====================================================================
        pnlTop = New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 114,
            .BackColor = Color.FromArgb(248, 250, 252),
            .Padding = New Padding(14, 8, 14, 6)
        }

        Dim lblTitulo As New Label() With {
            .Text = "ASIGNAR PADRES - Numeración Sistemática por Subensamblaje",
            .Font = New Font("Segoe UI", 10.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(15, 23, 42),
            .Location = New Point(12, 6),
            .AutoSize = True
        }
        pnlTop.Controls.Add(lblTitulo)

        Dim docNombre As String = If(_asmDoc IsNot Nothing, IOPath.GetFileName(_asmDoc.FullFileName), "Ensamblaje Activo")
        Dim rootCode As String = If(_asmDoc IsNot Nothing, ObtenerCodigoBaseEnsamble(_asmDoc), "-")
        Dim lblSubtitulo As New Label() With {
            .Text = String.Format("Documento: {0}   |   Código Raíz: {1}", docNombre, rootCode),
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(2, 132, 199),
            .Location = New Point(12, 28),
            .AutoSize = True
        }
        pnlTop.Controls.Add(lblSubtitulo)

        ' Acciones superiores derechas: Mantener al frente, Marcar, Desmarcar, Renumerar
        chkTopMost = New CheckBox() With {
            .Text = "📌 Al frente",
            .Checked = True,
            .Location = New Point(pnlTop.Width - 485, 8),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .AutoSize = True,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(71, 85, 105),
            .Cursor = Cursors.Hand
        }
        AddHandler chkTopMost.CheckedChanged, Sub() Me.TopMost = chkTopMost.Checked
        pnlTop.Controls.Add(chkTopMost)

        btnMarcarTodos = New Button() With {
            .Text = "☑ Marcar Todos",
            .Size = New Size(115, 26),
            .Location = New Point(pnlTop.Width - 375, 8),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .BackColor = Color.White,
            .ForeColor = Color.FromArgb(30, 41, 59),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 8.5F),
            .Cursor = Cursors.Hand
        }
        btnMarcarTodos.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        AddHandler btnMarcarTodos.Click, AddressOf BtnMarcarTodos_Click
        pnlTop.Controls.Add(btnMarcarTodos)

        btnDesmarcarTodos = New Button() With {
            .Text = "☐ Desmarcar",
            .Size = New Size(100, 26),
            .Location = New Point(pnlTop.Width - 252, 8),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .BackColor = Color.White,
            .ForeColor = Color.FromArgb(30, 41, 59),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 8.5F),
            .Cursor = Cursors.Hand
        }
        btnDesmarcarTodos.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        AddHandler btnDesmarcarTodos.Click, AddressOf BtnDesmarcarTodos_Click
        pnlTop.Controls.Add(btnDesmarcarTodos)

        btnAutoNumerar = New Button() With {
            .Text = "🔄 Renumerar",
            .Size = New Size(115, 26),
            .Location = New Point(pnlTop.Width - 144, 8),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .BackColor = Color.FromArgb(241, 245, 249),
            .ForeColor = Color.FromArgb(2, 132, 199),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .Cursor = Cursors.Hand
        }
        btnAutoNumerar.FlatAppearance.BorderColor = Color.FromArgb(2, 132, 199)
        AddHandler btnAutoNumerar.Click, AddressOf BtnAutoNumerar_Click
        pnlTop.Controls.Add(btnAutoNumerar)

        ' Fila 1 de opciones: Esquema, Dígitos, Separador, Filtro
        Dim lblEsquema As New Label() With {
            .Text = "Esquema:",
            .Location = New Point(12, 54),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(71, 85, 105)
        }
        pnlTop.Controls.Add(lblEsquema)

        cmbEsquema = New ComboBox() With {
            .Location = New Point(74, 50),
            .Size = New Size(130, 24),
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .BackColor = Color.White,
            .Font = New Font("Segoe UI", 8.5F)
        }
        cmbEsquema.Items.AddRange(New Object() {"[PADRE].[POS]", "[PADRE]-[POS]", "[PADRE]_[POS]"})
        cmbEsquema.SelectedIndex = 0
        AddHandler cmbEsquema.SelectedIndexChanged, AddressOf CmbEsquema_SelectedIndexChanged
        pnlTop.Controls.Add(cmbEsquema)

        Dim lblDigitos As New Label() With {
            .Text = "Dígitos:",
            .Location = New Point(216, 54),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(71, 85, 105)
        }
        pnlTop.Controls.Add(lblDigitos)

        cmbDigitos = New ComboBox() With {
            .Location = New Point(268, 50),
            .Size = New Size(80, 24),
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .BackColor = Color.White,
            .Font = New Font("Segoe UI", 8.5F)
        }
        cmbDigitos.Items.AddRange(New Object() {"01 (2 d.)", "001 (3 d.)", "1 (1 d.)"})
        cmbDigitos.SelectedIndex = 0
        AddHandler cmbDigitos.SelectedIndexChanged, Sub() AutoNumerarElementos()
        pnlTop.Controls.Add(cmbDigitos)

        Dim lblSep As New Label() With {
            .Text = "Sep.:",
            .Location = New Point(360, 54),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(71, 85, 105)
        }
        pnlTop.Controls.Add(lblSep)

        cmbSeparador = New ComboBox() With {
            .Location = New Point(394, 50),
            .Size = New Size(55, 24),
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .BackColor = Color.White,
            .Font = New Font("Segoe UI", 8.5F)
        }
        cmbSeparador.Items.AddRange(New Object() {".", "-", "_"})
        cmbSeparador.SelectedIndex = 0
        AddHandler cmbSeparador.SelectedIndexChanged, Sub() AutoNumerarElementos()
        pnlTop.Controls.Add(cmbSeparador)

        chkOmitirComerciales = New CheckBox() With {
            .Text = "Omitir Comerciales (COM)",
            .Location = New Point(460, 52),
            .AutoSize = True,
            .Checked = True,
            .Font = New Font("Segoe UI", 8.5F),
            .ForeColor = Color.FromArgb(30, 41, 59)
        }
        Dim tipCom As New ToolTip()
        tipCom.SetToolTip(chkOmitirComerciales, "Omite componentes comerciales (propiedad COM = '✓') para que mantengan su código comercial y no reciban numeración secuencial de padres.")
        AddHandler chkOmitirComerciales.CheckedChanged, Sub()
                                                            For Each it In _items
                                                                If it.IsCommercial Then
                                                                    it.Included = Not chkOmitirComerciales.Checked
                                                                End If
                                                            Next
                                                            AutoNumerarElementos()
                                                            ConstruirArbolJerarquico()
                                                            HabilitarVentanaInventor()
                                                        End Sub
        pnlTop.Controls.Add(chkOmitirComerciales)

        chkOmitirContentCenter = New CheckBox() With {
            .Text = "Omitir Content Center (CC)",
            .Location = New Point(636, 52),
            .AutoSize = True,
            .Checked = True,
            .Font = New Font("Segoe UI", 8.5F),
            .ForeColor = Color.FromArgb(30, 41, 59)
        }
        tipCom.SetToolTip(chkOmitirContentCenter, "Omite tornillería y piezas estándar de Content Center.")
        AddHandler chkOmitirContentCenter.CheckedChanged, Sub()
                                                              For Each it In _items
                                                                  If it.IsContentCenter Then
                                                                      it.Included = Not chkOmitirContentCenter.Checked
                                                                  End If
                                                              Next
                                                              AutoNumerarElementos()
                                                              ConstruirArbolJerarquico()
                                                              HabilitarVentanaInventor()
                                                          End Sub
        pnlTop.Controls.Add(chkOmitirContentCenter)

        Dim lblBuscar As New Label() With {
            .Text = "Buscar:",
            .Location = New Point(pnlTop.Width - 325, 54),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .AutoSize = True,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(71, 85, 105)
        }
        pnlTop.Controls.Add(lblBuscar)

        txtBuscar = New WinTextBox() With {
            .Location = New Point(pnlTop.Width - 265, 50),
            .Size = New Size(236, 24),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .BackColor = Color.White,
            .Font = New Font("Segoe UI", 8.5F)
        }
        AddHandler txtBuscar.TextChanged, AddressOf TxtBuscar_TextChanged
        pnlTop.Controls.Add(txtBuscar)

        ' Fila 2 de opciones: Sincronización de iProperties
        Dim lblProps As New Label() With {
            .Text = "Sincronizar en archivos:",
            .Location = New Point(12, 84),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(71, 85, 105)
        }
        pnlTop.Controls.Add(lblProps)

        chkActualizarStockNumber = New CheckBox() With {
            .Text = "Nº Almacenamiento (Stock)",
            .Location = New Point(160, 82),
            .AutoSize = True,
            .Checked = True,
            .Font = New Font("Segoe UI", 8.5F),
            .ForeColor = Color.FromArgb(30, 41, 59)
        }
        pnlTop.Controls.Add(chkActualizarStockNumber)

        chkActualizarPartNumber = New CheckBox() With {
            .Text = "Nº Pieza (Part Number)",
            .Location = New Point(340, 82),
            .AutoSize = True,
            .Checked = False,
            .Font = New Font("Segoe UI", 8.5F),
            .ForeColor = Color.FromArgb(30, 41, 59)
        }
        pnlTop.Controls.Add(chkActualizarPartNumber)

        chkActualizarPadre = New CheckBox() With {
            .Text = "Prop. PADRE",
            .Location = New Point(500, 82),
            .AutoSize = True,
            .Checked = True,
            .Font = New Font("Segoe UI", 8.5F),
            .ForeColor = Color.FromArgb(30, 41, 59)
        }
        pnlTop.Controls.Add(chkActualizarPadre)

        chkActualizarPos = New CheckBox() With {
            .Text = "Prop. POS",
            .Location = New Point(600, 82),
            .AutoSize = True,
            .Checked = True,
            .Font = New Font("Segoe UI", 8.5F),
            .ForeColor = Color.FromArgb(30, 41, 59)
        }
        pnlTop.Controls.Add(chkActualizarPos)

        ' =====================================================================
        ' 2. PANEL DERECHO: Detalle del componente, miniatura 3D y acciones
        ' =====================================================================
        pnlRight = New Panel() With {
            .Dock = DockStyle.Right,
            .Width = 265,
            .Padding = New Padding(10),
            .AutoScroll = True,
            .BackColor = Color.FromArgb(241, 245, 249)
        }

        Dim grpPreview As New GroupBox() With {
            .Text = "Detalle del Componente",
            .Dock = DockStyle.Fill,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(30, 41, 59)
        }

        picPreview = New PictureBox() With {
            .Location = New Point(15, 22),
            .Size = New Size(215, 150),
            .SizeMode = PictureBoxSizeMode.Zoom,
            .BorderStyle = BorderStyle.FixedSingle,
            .BackColor = Color.White
        }
        grpPreview.Controls.Add(picPreview)

        lblDetalleNombre = New Label() With {
            .Location = New Point(15, 178),
            .Size = New Size(215, 32),
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(15, 23, 42),
            .Text = "Componente: -"
        }
        grpPreview.Controls.Add(lblDetalleNombre)

        lblDetalleTipo = New Label() With {
            .Location = New Point(15, 212),
            .Size = New Size(215, 18),
            .Font = New Font("Segoe UI", 8.0F),
            .ForeColor = Color.FromArgb(71, 85, 105),
            .Text = "Tipo: -"
        }
        grpPreview.Controls.Add(lblDetalleTipo)

        lblDetallePadre = New Label() With {
            .Location = New Point(15, 232),
            .Size = New Size(215, 18),
            .Font = New Font("Segoe UI", 8.0F),
            .ForeColor = Color.FromArgb(71, 85, 105),
            .Text = "Padre: -"
        }
        grpPreview.Controls.Add(lblDetallePadre)

        lblDetalleStockActual = New Label() With {
            .Location = New Point(15, 252),
            .Size = New Size(215, 20),
            .Font = New Font("Segoe UI", 8.0F),
            .ForeColor = Color.FromArgb(71, 85, 105),
            .Text = "Stock Actual: -"
        }
        grpPreview.Controls.Add(lblDetalleStockActual)

        lblDetallePartNumber = New Label() With {
            .Location = New Point(15, 272),
            .Size = New Size(215, 20),
            .Font = New Font("Segoe UI", 8.0F),
            .ForeColor = Color.FromArgb(71, 85, 105),
            .Text = "Part Number: -"
        }
        grpPreview.Controls.Add(lblDetallePartNumber)

        lblDetallePos = New Label() With {
            .Location = New Point(15, 294),
            .Size = New Size(215, 20),
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(2, 132, 199),
            .Text = "Posición: -"
        }
        grpPreview.Controls.Add(lblDetallePos)

        lblDetallePropuesto = New Label() With {
            .Location = New Point(15, 316),
            .Size = New Size(215, 32),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(16, 185, 129),
            .Text = "Código Nuevo: -"
        }
        grpPreview.Controls.Add(lblDetallePropuesto)

        btnZoom = New Button() With {
            .Text = "🔍 Hacer Zoom en 3D",
            .Location = New Point(15, 352),
            .Size = New Size(215, 32),
            .BackColor = Color.White,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(30, 41, 59),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        btnZoom.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        AddHandler btnZoom.Click, AddressOf BtnZoom_Click
        grpPreview.Controls.Add(btnZoom)

        btnAbrir = New Button() With {
            .Text = "📂 Abrir Elemento",
            .Location = New Point(15, 388),
            .Size = New Size(215, 32),
            .BackColor = Color.White,
            .Font = New Font("Segoe UI", 8.5F),
            .ForeColor = Color.FromArgb(30, 41, 59),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        btnAbrir.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        AddHandler btnAbrir.Click, AddressOf BtnAbrir_Click
        grpPreview.Controls.Add(btnAbrir)

        Dim lblTip As New Label() With {
            .Text = "💡 Tip: Doble clic en cualquier nodo o fila para centrar y rotar libremente en la vista 3D.",
            .Location = New Point(15, 428),
            .Size = New Size(215, 42),
            .Font = New Font("Segoe UI", 7.5F, FontStyle.Italic),
            .ForeColor = Color.FromArgb(148, 163, 184)
        }
        grpPreview.Controls.Add(lblTip)

        btnAceptarLateral = New Button() With {
            .Text = "✔️ Aceptar y Aplicar",
            .Location = New Point(15, 480),
            .Size = New Size(215, 34),
            .BackColor = Color.FromArgb(2, 132, 199),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        btnAceptarLateral.FlatAppearance.BorderSize = 0
        AddHandler btnAceptarLateral.Click, AddressOf BtnAplicar_Click
        grpPreview.Controls.Add(btnAceptarLateral)

        btnCancelarLateral = New Button() With {
            .Text = "✖️ Cancelar",
            .Location = New Point(15, 520),
            .Size = New Size(215, 30),
            .BackColor = Color.White,
            .ForeColor = Color.FromArgb(51, 65, 85),
            .Font = New Font("Segoe UI", 8.5F),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        btnCancelarLateral.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        AddHandler btnCancelarLateral.Click, Sub() Me.Close()
        grpPreview.Controls.Add(btnCancelarLateral)

        pnlRight.Controls.Add(grpPreview)

        ' =====================================================================
        ' 3. PANEL INFERIOR: Resumen numérico y botones principales
        ' =====================================================================
        pnlBottom = New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 56,
            .Padding = New Padding(16, 8, 16, 8),
            .BackColor = Color.FromArgb(241, 245, 249)
        }

        Dim sepBottom As New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 1,
            .BackColor = Color.FromArgb(203, 213, 225)
        }
        pnlBottom.Controls.Add(sepBottom)

        lblResumen = New Label() With {
            .Location = New Point(16, 18),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(51, 65, 85),
            .Text = "Cargando estructura jerárquica del ensamble..."
        }
        pnlBottom.Controls.Add(lblResumen)

        Dim pnlBotonesInferiores As New FlowLayoutPanel() With {
            .Dock = DockStyle.Right,
            .FlowDirection = FlowDirection.RightToLeft,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .WrapContents = False,
            .Padding = New Padding(0, 10, 16, 8)
        }

        btnCancelar = New Button() With {
            .Text = "Cancelar",
            .Size = New Size(105, 34),
            .BackColor = Color.White,
            .ForeColor = Color.FromArgb(51, 65, 85),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand,
            .Margin = New Padding(10, 0, 0, 0)
        }
        btnCancelar.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        AddHandler btnCancelar.Click, Sub() Me.Close()

        btnAplicar = New Button() With {
            .Text = "Aceptar y Aplicar",
            .Size = New Size(180, 34),
            .BackColor = Color.FromArgb(2, 132, 199),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand,
            .Margin = New Padding(0, 0, 0, 0)
        }
        btnAplicar.FlatAppearance.BorderSize = 0
        AddHandler btnAplicar.Click, AddressOf BtnAplicar_Click

        pnlBotonesInferiores.Controls.Add(btnCancelar)
        pnlBotonesInferiores.Controls.Add(btnAplicar)
        pnlBottom.Controls.Add(pnlBotonesInferiores)

        AddHandler pnlBottom.Resize, Sub(s, e)
                                         lblResumen.Top = (pnlBottom.ClientSize.Height - lblResumen.Height) \ 2
                                     End Sub

        ' =====================================================================
        ' 4. CENTRO: SplitContainer con Árbol de Ensamblaje (Izquierda) y Grilla (Derecha)
        ' =====================================================================
        splitCenter = New SplitContainer() With {
            .Dock = DockStyle.Fill,
            .Orientation = Orientation.Vertical,
            .SplitterDistance = 290,
            .SplitterWidth = 6,
            .BackColor = Color.FromArgb(226, 232, 240)
        }

        ' --- PANEL IZQUIERDO DEL SPLIT: ÁRBOL JERÁRQUICO CON RAMAS Y CARPETAS ---
        Dim pnlTreeHost As New Panel() With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.White,
            .Padding = New Padding(0)
        }

        Dim pnlTreeHeader As New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 64,
            .BackColor = Color.FromArgb(248, 250, 252),
            .Padding = New Padding(8, 6, 8, 4)
        }

        Dim lblTreeTitle As New Label() With {
            .Text = "🌳 Estructura de Ensamblaje",
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(30, 41, 59),
            .Location = New Point(8, 6),
            .AutoSize = True
        }
        pnlTreeHeader.Controls.Add(lblTreeTitle)

        btnTreeExpandAll = New Button() With {
            .Text = "➕",
            .Size = New Size(30, 24),
            .Location = New Point(185, 4),
            .BackColor = Color.White,
            .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        btnTreeExpandAll.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        Dim tipExp As New ToolTip()
        tipExp.SetToolTip(btnTreeExpandAll, "Expandir todos los subensamblajes")
        AddHandler btnTreeExpandAll.Click, Sub()
                                               treeEstructura.ExpandAll()
                                               HabilitarVentanaInventor()
                                           End Sub
        pnlTreeHeader.Controls.Add(btnTreeExpandAll)

        btnTreeCollapseAll = New Button() With {
            .Text = "➖",
            .Size = New Size(30, 24),
            .Location = New Point(218, 4),
            .BackColor = Color.White,
            .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        btnTreeCollapseAll.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        tipExp.SetToolTip(btnTreeCollapseAll, "Contraer subensamblajes")
        AddHandler btnTreeCollapseAll.Click, Sub()
                                                 treeEstructura.CollapseAll()
                                                 If treeEstructura.Nodes.Count > 0 Then treeEstructura.Nodes(0).Expand()
                                                 HabilitarVentanaInventor()
                                             End Sub
        pnlTreeHeader.Controls.Add(btnTreeCollapseAll)

        btnTreeReset = New Button() With {
            .Text = "🌐",
            .Size = New Size(30, 24),
            .Location = New Point(251, 4),
            .BackColor = Color.White,
            .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        btnTreeReset.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        tipExp.SetToolTip(btnTreeReset, "Ver ensamble completo en la tabla")
        AddHandler btnTreeReset.Click, Sub()
                                           _filtroRamaSeleccionada = Nothing
                                           If treeEstructura.Nodes.Count > 0 Then treeEstructura.SelectedNode = treeEstructura.Nodes(0)
                                           RefrescarGrilla()
                                           HabilitarVentanaInventor()
                                       End Sub
        pnlTreeHeader.Controls.Add(btnTreeReset)

        chkFiltrarPorArbol = New CheckBox() With {
            .Text = "Filtrar tabla por rama seleccionada",
            .Location = New Point(8, 36),
            .AutoSize = True,
            .Checked = False,
            .Font = New Font("Segoe UI", 8.0F),
            .ForeColor = Color.FromArgb(71, 85, 105),
            .Cursor = Cursors.Hand
        }
        AddHandler chkFiltrarPorArbol.CheckedChanged, Sub()
                                                          If chkFiltrarPorArbol.Checked Then
                                                              _filtroRamaSeleccionada = treeEstructura.SelectedNode
                                                          Else
                                                              _filtroRamaSeleccionada = Nothing
                                                          End If
                                                          RefrescarGrilla()
                                                          HabilitarVentanaInventor()
                                                      End Sub
        pnlTreeHeader.Controls.Add(chkFiltrarPorArbol)

        treeEstructura = New TreeView() With {
            .Dock = DockStyle.Fill,
            .BorderStyle = BorderStyle.None,
            .Font = New Font("Segoe UI", 8.5F),
            .ForeColor = Color.FromArgb(15, 23, 42),
            .BackColor = Color.White,
            .FullRowSelect = True,
            .HideSelection = False,
            .ShowLines = True,
            .ShowPlusMinus = True,
            .ShowRootLines = True
        }
        AddHandler treeEstructura.AfterSelect, AddressOf TreeEstructura_AfterSelect
        AddHandler treeEstructura.NodeMouseDoubleClick, AddressOf TreeEstructura_NodeMouseDoubleClick

        pnlTreeHost.Controls.Add(treeEstructura)
        pnlTreeHost.Controls.Add(pnlTreeHeader)
        splitCenter.Panel1.Controls.Add(pnlTreeHost)

        ' --- PANEL DERECHO DEL SPLIT: DataGridView enriquecida ---
        _grid = New DataGridView() With {
            .Dock = DockStyle.Fill,
            .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.None,
            .GridColor = Color.FromArgb(226, 232, 240),
            .ForeColor = Color.FromArgb(15, 23, 42),
            .EnableHeadersVisualStyles = False,
            .RowHeadersVisible = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .MultiSelect = False,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .AllowUserToResizeRows = False,
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            .ColumnHeadersHeight = 32,
            .Font = New Font("Segoe UI", 8.5F),
            .ScrollBars = ScrollBars.Both
        }

        _grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(241, 245, 249)
        _grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(30, 41, 59)
        _grid.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        _grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        _grid.DefaultCellStyle.BackColor = Color.White
        _grid.DefaultCellStyle.ForeColor = Color.FromArgb(15, 23, 42)
        _grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(224, 242, 254)
        _grid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(15, 23, 42)
        _grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252)

        ConfigurarColumnasGrid()

        AddHandler _grid.SelectionChanged, AddressOf Grid_SelectionChanged
        AddHandler _grid.CellDoubleClick, AddressOf Grid_CellDoubleClick
        AddHandler _grid.CellValueChanged, AddressOf Grid_CellValueChanged
        AddHandler _grid.CurrentCellDirtyStateChanged, AddressOf Grid_CurrentCellDirtyStateChanged

        splitCenter.Panel2.Controls.Add(_grid)

        ' =====================================================================
        ' 5. ENSAMBLAJE DE CONTROLES: Orden de capas en Me.Controls
        ' =====================================================================
        Me.Controls.Add(splitCenter)
        Me.Controls.Add(pnlRight)
        Me.Controls.Add(pnlTop)
        Me.Controls.Add(pnlBottom)

        pnlBottom.BringToFront()
        pnlTop.BringToFront()
        pnlRight.BringToFront()
        splitCenter.BringToFront()

        Me.ResumeLayout(True)
    End Sub

    ''' <summary>
    ''' Configura las columnas del DataGridView con nombres claros e intuitivos.
    ''' </summary>
    Private Sub ConfigurarColumnasGrid()
        _grid.Columns.Clear()

        ' Columna Inc. (Inclusión)
        Dim colInc As New DataGridViewCheckBoxColumn() With {
            .Name = "ColInc",
            .HeaderText = "Inc.",
            .Width = 38,
            .SortMode = DataGridViewColumnSortMode.NotSortable
        }
        _grid.Columns.Add(colInc)

        ' Columna N° BOM
        Dim colItem As New DataGridViewTextBoxColumn() With {
            .Name = "ColItem",
            .HeaderText = "N°",
            .Width = 45,
            .ReadOnly = True,
            .DefaultCellStyle = New DataGridViewCellStyle() With {.Alignment = DataGridViewContentAlignment.MiddleCenter}
        }
        _grid.Columns.Add(colItem)

        ' Columna Nivel
        Dim colNivel As New DataGridViewTextBoxColumn() With {
            .Name = "ColNivel",
            .HeaderText = "Niv.",
            .Width = 42,
            .ReadOnly = True,
            .DefaultCellStyle = New DataGridViewCellStyle() With {.Alignment = DataGridViewContentAlignment.MiddleCenter}
        }
        _grid.Columns.Add(colNivel)

        ' Columna Componente / Archivo (con formato de rama 📁/📄)
        Dim colArchivo As New DataGridViewTextBoxColumn() With {
            .Name = "ColArchivo",
            .HeaderText = "Componente / Rama Jerárquica",
            .Width = 240,
            .ReadOnly = True,
            .DefaultCellStyle = New DataGridViewCellStyle() With {.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)}
        }
        _grid.Columns.Add(colArchivo)

        ' Columna Tipo
        Dim colTipo As New DataGridViewTextBoxColumn() With {
            .Name = "ColTipo",
            .HeaderText = "Tipo",
            .Width = 110,
            .ReadOnly = True
        }
        _grid.Columns.Add(colTipo)

        ' Columna Cantidad
        Dim colCant As New DataGridViewTextBoxColumn() With {
            .Name = "ColCant",
            .HeaderText = "Cant.",
            .Width = 50,
            .ReadOnly = True,
            .DefaultCellStyle = New DataGridViewCellStyle() With {.Alignment = DataGridViewContentAlignment.MiddleCenter}
        }
        _grid.Columns.Add(colCant)

        ' Columna Ensamble Padre
        Dim colPadre As New DataGridViewTextBoxColumn() With {
            .Name = "ColPadre",
            .HeaderText = "Ensamble Padre",
            .Width = 140,
            .ReadOnly = True
        }
        _grid.Columns.Add(colPadre)

        ' Columna POS (Posición Secuencial)
        Dim colPos As New DataGridViewTextBoxColumn() With {
            .Name = "ColPos",
            .HeaderText = "POS",
            .Width = 55,
            .ReadOnly = False,
            .DefaultCellStyle = New DataGridViewCellStyle() With {
                .Alignment = DataGridViewContentAlignment.MiddleCenter,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
                .ForeColor = Color.FromArgb(2, 132, 199)
            }
        }
        _grid.Columns.Add(colPos)

        ' Columna Nº Almacenamiento Actual (Stock Number)
        Dim colStock As New DataGridViewTextBoxColumn() With {
            .Name = "ColStockActual",
            .HeaderText = "Nº Almacenamiento Actual",
            .Width = 150,
            .ReadOnly = True
        }
        _grid.Columns.Add(colStock)

        ' Columna Nuevo Código Propuesto
        Dim colNuevoCodigo As New DataGridViewTextBoxColumn() With {
            .Name = "ColNuevoCodigo",
            .HeaderText = "Nuevo Código Propuesto",
            .Width = 175,
            .ReadOnly = False,
            .DefaultCellStyle = New DataGridViewCellStyle() With {
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
                .ForeColor = Color.FromArgb(16, 185, 129)
            }
        }
        _grid.Columns.Add(colNuevoCodigo)

        ' Columna Part Number Actual
        Dim colPartNumber As New DataGridViewTextBoxColumn() With {
            .Name = "ColPartNumber",
            .HeaderText = "Part Number Actual",
            .Width = 140,
            .ReadOnly = True
        }
        _grid.Columns.Add(colPartNumber)
    End Sub

    ''' <summary>
    ''' Carga los componentes priorizando la Lista de Materiales (BOM) estructurada de Inventor
    ''' para asegurar que se listen exactamente las mismas piezas y cantidades que ve el usuario en el BOM.
    ''' Si el BOM no está habilitado o tiene 0 filas, ejecuta un recorrido recursivo de Ocurrencias como respaldo seguro.
    ''' </summary>
    Private Sub CargarJerarquiaDesdeBOM()
        _items.Clear()
        If _asmDoc Is Nothing Then Return

        Dim rootCode As String = ObtenerCodigoBaseEnsamble(_asmDoc)
        Dim cargoDesdeBOM As Boolean = False

        ' 1. Intentar recorrer BOM estructurado de Inventor
        Try
            Dim oBOM As BOM = _asmDoc.BOM
            Try
                oBOM.StructuredViewEnabled = True
            Catch
            End Try
            Try
                oBOM.StructuredViewFirstLevelOnly = False
            Catch
            End Try

            Dim structView As BOMView = Nothing
            If oBOM.BOMViews IsNot Nothing Then
                For Each bv As BOMView In oBOM.BOMViews
                    If bv.ViewType = BOMViewTypeEnum.kStructuredBOMViewType OrElse
                       bv.Name.IndexOf("Structured", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       bv.Name.IndexOf("Estructurad", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        structView = bv
                        Exit For
                    End If
                Next
            End If

            If structView IsNot Nothing AndAlso structView.BOMRows IsNot Nothing AndAlso structView.BOMRows.Count > 0 Then
                RecorrerBOMRows(structView.BOMRows, rootCode, 1)
                If _items.Count > 0 Then
                    cargoDesdeBOM = True
                End If
            End If
        Catch
            cargoDesdeBOM = False
        End Try

        ' 2. Fallback robusto: si el BOM estructurado no arrojó filas, recorrer Occurrences directamente
        If Not cargoDesdeBOM OrElse _items.Count = 0 Then
            Try
                RecorrerOccurrencesFallback(_asmDoc.ComponentDefinition.Occurrences, rootCode, 1)
            Catch
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Recorre recursivamente las filas de la vista estructurada del BOM de Inventor
    ''' </summary>
    Private Sub RecorrerBOMRows(rows As BOMRowsEnumerator, padreCode As String, nivel As Integer)
        If rows Is Nothing Then Return

        For Each row As BOMRow In rows
            If row Is Nothing Then Continue For

            Dim compDef As ComponentDefinition = Nothing
            Try
                If row.ComponentDefinitions IsNot Nothing AndAlso row.ComponentDefinitions.Count > 0 Then
                    compDef = row.ComponentDefinitions.Item(1)
                End If
            Catch
            End Try

            If compDef Is Nothing Then Continue For

            Dim doc As Document = Nothing
            Try : doc = compDef.Document : Catch : End Try
            If doc Is Nothing Then Continue For

            Dim isAsm As Boolean = (doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject)
            Dim isCC As Boolean = False
            Try
                If compDef.IsContentMember Then isCC = True
            Catch
            End Try
            If Not isCC Then
                Try : isCC = (doc.PropertySets.Item("Content Library Component Properties").Count > 0) : Catch : End Try
            End If

            Dim structStr As String = "Normal"
            Try
                Select Case row.BOMStructure
                    Case BOMStructureEnum.kPurchasedBOMStructure : structStr = "Comprado"
                    Case BOMStructureEnum.kInseparableBOMStructure : structStr = "Inseparable"
                    Case BOMStructureEnum.kPhantomBOMStructure : structStr = "Fantasma"
                    Case Else : structStr = "Normal"
                End Select
            Catch
            End Try

            Dim isCom As Boolean = EsComponenteComercial(doc, row.BOMStructure)
            Dim stockNum As String = ObtenerStockNumber(doc)
            Dim partNum As String = ObtenerPartNumber(doc)
            Dim categoriaFab As String = If(isCom, "COM", ObtenerCategoriaFabricacion(doc, isAsm))

            Dim item As New ElementoPadreItem() With {
                .Document = doc,
                .DocumentPath = doc.FullFileName,
                .FileName = IOPath.GetFileName(doc.FullFileName),
                .Extension = If(isAsm, ".iam", ".ipt"),
                .ItemNumber = row.ItemNumber,
                .Quantity = row.ItemQuantity,
                .BomStructure = structStr,
                .IsAssembly = isAsm,
                .IsCommercial = isCom,
                .IsContentCenter = isCC,
                .Categoria = categoriaFab,
                .CurrentPadre = ObtenerCustomProp(doc, "PADRE"),
                .SuggestedPadre = padreCode,
                .CurrentPartNumber = partNum,
                .CurrentStockNumber = stockNum,
                .Level = nivel,
                .Included = (Not isCom AndAlso Not isCC),
                .ProposedPosition = If(isCom OrElse isCC, "-", ""),
                .ProposedCode = If(isCom, If(Not String.IsNullOrEmpty(stockNum), stockNum, If(Not String.IsNullOrEmpty(partNum), partNum, IOPath.GetFileNameWithoutExtension(doc.FullFileName))), "")
            }

            ' Encontrar ocurrencia real en el ensamblaje para resaltado 3D y zoom
            Try
                item.Occurrence = BuscarOcurrenciaPorDocumento(_asmDoc.ComponentDefinition.Occurrences, doc.FullFileName)
            Catch
            End Try

            _items.Add(item)

            ' Si es subensamblaje de fabricación (NO comercial) y tiene filas hijas, descender en jerarquía recursivamente
            ' Los componentes comerciales se adquieren terminados y nunca reciben renumeración interna de padres
            If isAsm AndAlso Not isCom Then
                Dim subPadre As String = If(Not String.IsNullOrEmpty(item.CurrentStockNumber), item.CurrentStockNumber,
                                         If(Not String.IsNullOrEmpty(item.CurrentPartNumber), item.CurrentPartNumber, item.FileName))
                Try
                    If row.ChildRows IsNot Nothing AndAlso row.ChildRows.Count > 0 Then
                        RecorrerBOMRows(row.ChildRows, subPadre, nivel + 1)
                    End If
                Catch
                End Try
            End If
        Next
    End Sub

    ''' <summary>
    ''' Recorrido alternativo usando Occurrences cuando el BOM no está activo
    ''' </summary>
    Private Sub RecorrerOccurrencesFallback(occs As ComponentOccurrences, padreCode As String, nivel As Integer)
        If occs Is Nothing Then Return

        For Each occ As ComponentOccurrence In occs
            If occ Is Nothing OrElse occ.Suppressed Then Continue For

            Dim doc As Document = Nothing
            Try : doc = occ.Definition.Document : Catch : End Try
            If doc Is Nothing Then Continue For

            Dim docPath As String = doc.FullFileName

            ' Si ya se incluyó este documento en este nivel de jerarquía, incrementar cantidad
            Dim existente = _items.Find(Function(i) String.Equals(i.DocumentPath, docPath, StringComparison.OrdinalIgnoreCase) AndAlso i.Level = nivel)
            If existente IsNot Nothing Then
                existente.Quantity += 1
                Continue For
            End If

            Dim isAsm As Boolean = (occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject)
            Dim isCC As Boolean = False
            Try : isCC = occ.Definition.IsContentMember : Catch : End Try
            If Not isCC Then
                Try : isCC = (doc.PropertySets.Item("Content Library Component Properties").Count > 0) : Catch : End Try
            End If

            Dim isCom As Boolean = EsComponenteComercial(doc, BOMStructureEnum.kNormalBOMStructure)
            Dim stockNum As String = ObtenerStockNumber(doc)
            Dim partNum As String = ObtenerPartNumber(doc)
            Dim categoriaFab As String = If(isCom, "COM", ObtenerCategoriaFabricacion(doc, isAsm))

            Dim item As New ElementoPadreItem() With {
                .Occurrence = occ,
                .Document = doc,
                .DocumentPath = docPath,
                .FileName = IOPath.GetFileName(docPath),
                .Extension = If(isAsm, ".iam", ".ipt"),
                .ItemNumber = (_items.Count + 1).ToString(),
                .Quantity = 1,
                .BomStructure = If(isCom, "Comprado", "Normal"),
                .IsAssembly = isAsm,
                .IsCommercial = isCom,
                .IsContentCenter = isCC,
                .Categoria = categoriaFab,
                .CurrentPadre = ObtenerCustomProp(doc, "PADRE"),
                .SuggestedPadre = padreCode,
                .CurrentPartNumber = partNum,
                .CurrentStockNumber = stockNum,
                .Level = nivel,
                .Included = (Not isCom AndAlso Not isCC),
                .ProposedPosition = If(isCom OrElse isCC, "-", ""),
                .ProposedCode = If(isCom, If(Not String.IsNullOrEmpty(stockNum), stockNum, If(Not String.IsNullOrEmpty(partNum), partNum, IOPath.GetFileNameWithoutExtension(docPath))), "")
            }

            _items.Add(item)

            ' Descender en subensamblajes de fabricación únicamente
            If isAsm AndAlso Not isCom AndAlso occ.SubOccurrences IsNot Nothing AndAlso occ.SubOccurrences.Count > 0 Then
                Dim subPadre As String = If(Not String.IsNullOrEmpty(item.CurrentStockNumber), item.CurrentStockNumber,
                                         If(Not String.IsNullOrEmpty(item.CurrentPartNumber), item.CurrentPartNumber, item.FileName))
                RecorrerOccurrencesFallback(occ.SubOccurrences, subPadre, nivel + 1)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Detecta si un documento es un componente comercial según la propiedad COM = "✓",
    ''' o si su estructura en la lista de materiales es Comprado (Purchased).
    ''' </summary>
    Private Function EsComponenteComercial(doc As Document, Optional bomStruct As BOMStructureEnum = BOMStructureEnum.kNormalBOMStructure) As Boolean
        If doc Is Nothing Then Return False
        Try
            If bomStruct = BOMStructureEnum.kPurchasedBOMStructure Then Return True

            ' Verificar propiedad personalizada COM (case-insensitive, acepta ✓, SI, TRUE, etc.)
            Dim comVal As String = ObtenerCustomProp(doc, "COM")
            If Not String.IsNullOrEmpty(comVal) Then
                If comVal.Contains("✓") OrElse comVal.Contains("✔") OrElse comVal.Contains("√") OrElse
                   String.Equals(comVal, "SI", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(comVal, "S", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(comVal, "TRUE", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(comVal, "1", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(comVal, "X", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(comVal, "V", StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            End If
        Catch
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Determina la categoría de fabricación (CYP, MEC, MC, ARM) leyendo las iProperties del documento.
    ''' </summary>
    Private Function ObtenerCategoriaFabricacion(doc As Document, isAsm As Boolean) As String
        If doc Is Nothing Then Return If(isAsm, "ARM", "CYP")
        Try
            If ObtenerCustomProp(doc, "MEC") = "✓" Then Return "MEC"
            If ObtenerCustomProp(doc, "MC") = "✓" Then Return "MC"
            If ObtenerCustomProp(doc, "CYP") = "✓" Then Return "CYP"
            If ObtenerCustomProp(doc, "ARM") = "✓" Then Return "ARM"

            Dim catDirecta As String = ObtenerCustomProp(doc, "CATEGORIA")
            If Not String.IsNullOrEmpty(catDirecta) Then Return catDirecta

            Return If(isAsm, "ARM", "CYP")
        Catch
            Return If(isAsm, "ARM", "CYP")
        End Try
    End Function

    ''' <summary>
    ''' Construye la estructura de carpetas y ramas jerárquicas en el TreeView.
    ''' Agrupa las partes bajo su respectivo ensamble o subensamblaje con íconos de carpetas 📁 y piezas 📄.
    ''' Aísla los componentes comerciales (COM) en ramas dedicadas para que no interfieran con la numeración de los padres.
    ''' </summary>
    Private Sub ConstruirArbolJerarquico()
        treeEstructura.BeginUpdate()
        treeEstructura.Nodes.Clear()

        Dim rootCode As String = ObtenerCodigoBaseEnsamble(_asmDoc)
        Dim docNombre As String = If(_asmDoc IsNot Nothing, IOPath.GetFileName(_asmDoc.FullFileName), "Ensamblaje Principal")

        ' Nodo raíz del ensamblaje
        Dim rootNode As New TreeNode("📁 " & docNombre & " [" & rootCode & "]") With {
            .Tag = Nothing,
            .NodeFont = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        }
        treeEstructura.Nodes.Add(rootNode)

        ' Diccionario para encontrar rápidamente el nodo de cada ensamble padre
        Dim dictAsmNodes As New Dictionary(Of String, TreeNode)(StringComparer.OrdinalIgnoreCase)
        dictAsmNodes(rootCode) = rootNode
        dictAsmNodes(docNombre) = rootNode
        dictAsmNodes(IOPath.GetFileNameWithoutExtension(docNombre)) = rootNode

        ' Separar piezas de fabricación vs comerciales (COM) vs Content Center (CC)
        Dim itemsFabricacion As New List(Of ElementoPadreItem)()
        Dim itemsComerciales As New List(Of ElementoPadreItem)()
        Dim itemsCC As New List(Of ElementoPadreItem)()

        For Each item In _items
            If item.IsCommercial Then
                itemsComerciales.Add(item)
            ElseIf item.IsContentCenter Then
                itemsCC.Add(item)
            Else
                itemsFabricacion.Add(item)
            End If
        Next

        ' 1. Añadir componentes y subensamblajes de fabricación formando la estructura principal
        For Each item In itemsFabricacion
            Dim parentNode As TreeNode = rootNode
            If Not String.IsNullOrEmpty(item.SuggestedPadre) AndAlso dictAsmNodes.ContainsKey(item.SuggestedPadre) Then
                parentNode = dictAsmNodes(item.SuggestedPadre)
            End If

            Dim etiqueta As String = item.IconoPrefix & item.FileName
            If Not String.IsNullOrEmpty(item.ProposedCode) Then
                etiqueta &= " (" & item.ProposedCode & ")"
            ElseIf Not String.IsNullOrEmpty(item.CurrentStockNumber) Then
                etiqueta &= " (" & item.CurrentStockNumber & ")"
            End If
            If item.Quantity > 1 Then
                etiqueta &= " x" & item.Quantity
            End If

            Dim nodoItem As New TreeNode(etiqueta) With {
                .Tag = item
            }

            If item.IsAssembly Then
                nodoItem.NodeFont = New Font("Segoe UI", 8.5F, FontStyle.Bold)
                nodoItem.ForeColor = Color.FromArgb(2, 132, 199)

                ' Registrar el subensamblaje para que sus hijos cuelguen de él
                If Not String.IsNullOrEmpty(item.ProposedCode) Then dictAsmNodes(item.ProposedCode) = nodoItem
                If Not String.IsNullOrEmpty(item.CurrentStockNumber) Then dictAsmNodes(item.CurrentStockNumber) = nodoItem
                If Not String.IsNullOrEmpty(item.CurrentPartNumber) Then dictAsmNodes(item.CurrentPartNumber) = nodoItem
                dictAsmNodes(item.FileName) = nodoItem
                dictAsmNodes(IOPath.GetFileNameWithoutExtension(item.FileName)) = nodoItem
            End If

            parentNode.Nodes.Add(nodoItem)
        Next

        ' 2. Añadir componentes comerciales (COM) AISLADOS en carpetas dedicadas bajo su respectivo padre
        Dim comPorPadre As New Dictionary(Of String, List(Of ElementoPadreItem))(StringComparer.OrdinalIgnoreCase)
        For Each comItem In itemsComerciales
            Dim pKey As String = If(Not String.IsNullOrEmpty(comItem.SuggestedPadre), comItem.SuggestedPadre, rootCode)
            If Not comPorPadre.ContainsKey(pKey) Then
                comPorPadre(pKey) = New List(Of ElementoPadreItem)()
            End If
            comPorPadre(pKey).Add(comItem)
        Next

        For Each kvp In comPorPadre
            Dim pNode As TreeNode = rootNode
            If dictAsmNodes.ContainsKey(kvp.Key) Then
                pNode = dictAsmNodes(kvp.Key)
            End If

            Dim nodoCarpetaCom As New TreeNode("📦 Componentes Comerciales (COM) (" & kvp.Value.Count & ")") With {
                .Tag = Nothing,
                .NodeFont = New Font("Segoe UI", 8.0F, FontStyle.Bold),
                .ForeColor = Color.FromArgb(5, 150, 105)
            }
            pNode.Nodes.Add(nodoCarpetaCom)

            For Each comItem In kvp.Value
                Dim etiq As String = "📦 " & comItem.FileName
                Dim comCod As String = comItem.CodigoComercial
                If Not String.IsNullOrEmpty(comCod) Then
                    etiq &= " [" & comCod & "]"
                End If
                If comItem.Quantity > 1 Then
                    etiq &= " x" & comItem.Quantity
                End If

                Dim nodoCom As New TreeNode(etiq) With {
                    .Tag = comItem,
                    .ForeColor = Color.FromArgb(5, 150, 105)
                }
                nodoCarpetaCom.Nodes.Add(nodoCom)
            Next
        Next

        ' 3. Añadir componentes de Content Center (CC) aislados en carpetas dedicadas
        Dim ccPorPadre As New Dictionary(Of String, List(Of ElementoPadreItem))(StringComparer.OrdinalIgnoreCase)
        For Each ccItem In itemsCC
            Dim pKey As String = If(Not String.IsNullOrEmpty(ccItem.SuggestedPadre), ccItem.SuggestedPadre, rootCode)
            If Not ccPorPadre.ContainsKey(pKey) Then
                ccPorPadre(pKey) = New List(Of ElementoPadreItem)()
            End If
            ccPorPadre(pKey).Add(ccItem)
        Next

        For Each kvp In ccPorPadre
            Dim pNode As TreeNode = rootNode
            If dictAsmNodes.ContainsKey(kvp.Key) Then
                pNode = dictAsmNodes(kvp.Key)
            End If

            Dim nodoCarpetaCC As New TreeNode("🔩 Tornillería y Content Center (CC) (" & kvp.Value.Count & ")") With {
                .Tag = Nothing,
                .NodeFont = New Font("Segoe UI", 8.0F, FontStyle.Italic),
                .ForeColor = Color.FromArgb(100, 116, 139)
            }
            pNode.Nodes.Add(nodoCarpetaCC)

            For Each ccItem In kvp.Value
                Dim etiq As String = "🔩 " & ccItem.FileName
                If ccItem.Quantity > 1 Then etiq &= " x" & ccItem.Quantity
                Dim nodoCC As New TreeNode(etiq) With {
                    .Tag = ccItem,
                    .ForeColor = Color.FromArgb(100, 116, 139)
                }
                nodoCarpetaCC.Nodes.Add(nodoCC)
            Next
        Next

        rootNode.Expand()
        For Each child As TreeNode In rootNode.Nodes
            child.Expand()
        Next

        treeEstructura.EndUpdate()
    End Sub

    ''' <summary>
    ''' Busca recursivamente la primera ocurrencia de un documento en el árbol de ensamblaje.
    ''' </summary>
    Private Function BuscarOcurrenciaPorDocumento(occs As ComponentOccurrences, fullFileName As String) As ComponentOccurrence
        If occs Is Nothing OrElse String.IsNullOrEmpty(fullFileName) Then Return Nothing
        Try
            For Each occ As ComponentOccurrence In occs
                If occ.Suppressed Then Continue For
                Try
                    If occ.Definition IsNot Nothing AndAlso occ.Definition.Document IsNot Nothing Then
                        If String.Equals(occ.Definition.Document.FullFileName, fullFileName, StringComparison.OrdinalIgnoreCase) Then
                            Return occ
                        End If
                    End If
                Catch
                End Try

                If occ.SubOccurrences IsNot Nothing AndAlso occ.SubOccurrences.Count > 0 Then
                    Dim found As ComponentOccurrence = BuscarOcurrenciaPorDocumento(occ.SubOccurrences, fullFileName)
                    If found IsNot Nothing Then Return found
                End If
            Next
        Catch
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Detecta automáticamente si el ensamble utiliza punto '.' o guión '-' según sus códigos
    ''' </summary>
    Private Sub AutoDetectarSeparadorYPrefijo()
        Dim rootCode As String = ObtenerCodigoBaseEnsamble(_asmDoc)

        If rootCode.Contains(".") Then
            cmbSeparador.SelectedItem = "."
            cmbEsquema.SelectedIndex = 0 ' [PADRE].[POS]
        Else
            cmbSeparador.SelectedItem = "-"
            cmbEsquema.SelectedIndex = 1 ' [PADRE]-[POS]
        End If
    End Sub

    ''' <summary>
    ''' Calcula la numeración secuencial de cada pieza agrupada bajo su ensamble padre directo.
    ''' </summary>
    Private Sub AutoNumerarElementos()
        Dim parentCounters As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Dim digitos As Integer = If(cmbDigitos.SelectedIndex = 1, 3, If(cmbDigitos.SelectedIndex = 2, 1, 2))
        Dim separador As String = If(cmbSeparador.SelectedItem IsNot Nothing, cmbSeparador.SelectedItem.ToString(), ".")

        For Each item In _items
            ' 1. Comerciales (COM): NUNCA reciben numeración secuencial de padres.
            ' Conservan siempre su código comercial propio y posición '-'.
            If item.IsCommercial Then
                item.ProposedPosition = "-"
                item.ProposedCode = item.CodigoComercial
                If chkOmitirComerciales.Checked Then
                    item.Included = False
                End If
                Continue For
            End If

            ' 2. Content Center (CC): Se omiten de la numeración de piezas fabricadas.
            If item.IsContentCenter Then
                item.ProposedPosition = "-"
                item.ProposedCode = If(Not String.IsNullOrEmpty(item.CurrentStockNumber), item.CurrentStockNumber, item.CurrentPartNumber)
                If chkOmitirContentCenter.Checked Then
                    item.Included = False
                End If
                Continue For
            End If

            ' 3. Omitir si el usuario desmarcó el checkbox en la grilla
            If Not item.Included Then
                item.ProposedPosition = "-"
                item.ProposedCode = item.CurrentStockNumber
                Continue For
            End If

            ' 4. Respetar edición manual si el usuario modificó la celda
            If item.ManualOverride Then Continue For

            ' 5. Numeración secuencial para piezas y subensamblajes de fabricación
            Dim padreKey As String = If(Not String.IsNullOrEmpty(item.SuggestedPadre), item.SuggestedPadre, "PADRE")
            If Not parentCounters.ContainsKey(padreKey) Then
                parentCounters(padreKey) = 1
            End If

            Dim currentSeq As Integer = parentCounters(padreKey)
            parentCounters(padreKey) += 1

            Dim posStr As String = currentSeq.ToString().PadLeft(digitos, "0"c)
            item.ProposedPosition = posStr
            item.ProposedCode = padreKey & separador & posStr
        Next

        RefrescarGrilla()
    End Sub

    ''' <summary>
    ''' Refresca visualmente la grilla asegurando visibilidad clara de las ramas y colores impecables.
    ''' Muestra primero las piezas de fabricación, seguidas de los componentes comerciales (COM) y de librería (CC).
    ''' Permite filtrar por la rama seleccionada en el árbol si la opción está activa.
    ''' </summary>
    Private Sub RefrescarGrilla()
        _grid.SuspendLayout()
        _grid.Rows.Clear()

        Dim filtroTexto As String = If(txtBuscar IsNot Nothing, txtBuscar.Text.Trim().ToLower(), "")
        Dim aNumerarCount As Integer = 0
        Dim modificadosCount As Integer = 0

        ' Obtener conjunto de items permitidos según el filtro de rama del árbol
        Dim itemsFiltradosRama As HashSet(Of ElementoPadreItem) = Nothing
        If chkFiltrarPorArbol IsNot Nothing AndAlso chkFiltrarPorArbol.Checked AndAlso _filtroRamaSeleccionada IsNot Nothing Then
            itemsFiltradosRama = ObtenerItemsDeRamaNodo(_filtroRamaSeleccionada)
        End If

        ' Agrupar elementos para presentación ordenada: Fabricación primero, luego Comerciales, luego Content Center
        Dim itemsParaMostrar As New List(Of ElementoPadreItem)()
        itemsParaMostrar.AddRange(_items.Where(Function(i) Not i.IsCommercial AndAlso Not i.IsContentCenter))
        itemsParaMostrar.AddRange(_items.Where(Function(i) i.IsCommercial))
        itemsParaMostrar.AddRange(_items.Where(Function(i) i.IsContentCenter))

        For Each item In itemsParaMostrar
            ' Filtro de rama de árbol
            If itemsFiltradosRama IsNot Nothing AndAlso Not itemsFiltradosRama.Contains(item) Then
                Continue For
            End If

            ' Filtro de texto de búsqueda
            If Not String.IsNullOrEmpty(filtroTexto) Then
                Dim coincide As Boolean = item.FileName.ToLower().Contains(filtroTexto) OrElse
                                          item.CurrentStockNumber.ToLower().Contains(filtroTexto) OrElse
                                          item.CurrentPartNumber.ToLower().Contains(filtroTexto) OrElse
                                          item.SuggestedPadre.ToLower().Contains(filtroTexto) OrElse
                                          item.ProposedCode.ToLower().Contains(filtroTexto)
                If Not coincide Then Continue For
            End If

            Dim rowIdx As Integer = _grid.Rows.Add(
                item.Included,
                item.ItemNumber,
                "L" & item.Level,
                item.RamaFormateada,
                item.ColumnaTipo,
                item.Quantity,
                item.SuggestedPadre,
                item.ProposedPosition,
                item.CurrentStockNumber,
                item.ProposedCode,
                item.CurrentPartNumber
            )

            Dim row As DataGridViewRow = _grid.Rows(rowIdx)
            row.Tag = item

            ' Distinción visual para ensamblajes padres (subrayado suave de carpeta)
            If item.IsAssembly Then
                row.DefaultCellStyle.BackColor = Color.FromArgb(241, 245, 249)
                row.DefaultCellStyle.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
                row.Cells("ColTipo").Style.ForeColor = Color.FromArgb(2, 132, 199)
            ElseIf item.IsCommercial Then
                row.Cells("ColTipo").Style.ForeColor = Color.FromArgb(5, 150, 105) ' Emerald
            ElseIf item.IsContentCenter Then
                row.Cells("ColTipo").Style.ForeColor = Color.FromArgb(100, 116, 139)
            End If

            If item.TieneCambio Then
                row.Cells("ColNuevoCodigo").Style.ForeColor = Color.FromArgb(2, 132, 199)
                row.Cells("ColNuevoCodigo").Style.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
                modificadosCount += 1
            End If

            If item.Included AndAlso Not item.IsCommercial AndAlso Not item.IsContentCenter Then
                aNumerarCount += 1
            End If
        Next

        _grid.ResumeLayout()

        Dim ramaInfo As String = If(itemsFiltradosRama IsNot Nothing, " [Filtrado por rama: " & itemsFiltradosRama.Count & " items]", "")
        lblResumen.Text = String.Format("Total elementos: {0}   |   Piezas a numerar: {1}   |   Cambios proyectados: {2}{3}",
                                        _items.Count, aNumerarCount, modificadosCount, ramaInfo)

        ' Seleccionar primera fila si hay elementos disponibles
        If _grid.Rows.Count > 0 AndAlso _grid.CurrentRow Is Nothing Then
            _grid.CurrentCell = _grid.Rows(0).Cells(0)
        End If
    End Sub

    ''' <summary>
    ''' Obtiene recursivamente todos los items vinculados a un nodo del TreeView y sus descendientes.
    ''' </summary>
    Private Function ObtenerItemsDeRamaNodo(nodo As TreeNode) As HashSet(Of ElementoPadreItem)
        Dim res As New HashSet(Of ElementoPadreItem)()
        If nodo Is Nothing Then Return res

        ' Si es el nodo raíz (Tag = Nothing), incluir todos los items
        If nodo.Tag Is Nothing Then
            For Each it In _items
                res.Add(it)
            Next
            Return res
        End If

        Dim rootItem = TryCast(nodo.Tag, ElementoPadreItem)
        If rootItem IsNot Nothing Then res.Add(rootItem)

        RecorrerHijosNodo(nodo, res)
        Return res
    End Function

    Private Sub RecorrerHijosNodo(nodo As TreeNode, ByRef res As HashSet(Of ElementoPadreItem))
        For Each child As TreeNode In nodo.Nodes
            Dim it = TryCast(child.Tag, ElementoPadreItem)
            If it IsNot Nothing Then res.Add(it)
            RecorrerHijosNodo(child, res)
        Next
    End Sub

    ''' <summary>
    ''' Evento de selección en el árbol: sincroniza la selección con la tabla y actualiza detalles.
    ''' </summary>
    Private Sub TreeEstructura_AfterSelect(sender As Object, e As TreeViewEventArgs)
        If _isSyncingSelection OrElse e.Node Is Nothing Then Return

        _isSyncingSelection = True
        Try
            Dim item As ElementoPadreItem = TryCast(e.Node.Tag, ElementoPadreItem)

            If chkFiltrarPorArbol.Checked Then
                _filtroRamaSeleccionada = e.Node
                RefrescarGrilla()
            Else
                ' Localizar y seleccionar la fila correspondiente en la grilla
                If item IsNot Nothing Then
                    For Each row As DataGridViewRow In _grid.Rows
                        If row.Tag Is item Then
                            _grid.ClearSelection()
                            row.Selected = True
                            _grid.CurrentCell = row.Cells(0)
                            Exit For
                        End If
                    Next
                End If
            End If

            ActualizarDetallesItem(item)

            If item IsNot Nothing AndAlso item.Occurrence IsNot Nothing Then
                ResaltarElemento(item.Occurrence)
            End If

            HabilitarVentanaInventor()
        Finally
            _isSyncingSelection = False
        End Try
    End Sub

    ''' <summary>
    ''' Doble clic en el árbol: centra y resalta la pieza en Inventor 3D.
    ''' </summary>
    Private Sub TreeEstructura_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs)
        Dim item As ElementoPadreItem = TryCast(e.Node.Tag, ElementoPadreItem)
        If item IsNot Nothing Then
            HacerZoom3D(item)
            HabilitarVentanaInventor()
        End If
    End Sub

    ''' <summary>
    ''' Actualiza los detalles del panel lateral derecho para el item seleccionado.
    ''' </summary>
    Private Sub ActualizarDetallesItem(item As ElementoPadreItem)
        If item Is Nothing Then
            picPreview.Image = Nothing
            lblDetalleNombre.Text = "Componente: -"
            lblDetalleTipo.Text = "Tipo: -"
            lblDetallePadre.Text = "Padre: -"
            lblDetalleStockActual.Text = "Stock Actual: -"
            lblDetallePartNumber.Text = "Part Number: -"
            lblDetallePos.Text = "Posición asignada: -"
            lblDetallePropuesto.Text = "Código Nuevo: -"
            Return
        End If

        lblDetalleNombre.Text = item.FileName
        lblDetalleTipo.Text = "Tipo: " & item.ColumnaTipo
        lblDetallePadre.Text = "Padre: " & item.SuggestedPadre
        lblDetalleStockActual.Text = "Stock Actual: " & If(Not String.IsNullOrEmpty(item.CurrentStockNumber), item.CurrentStockNumber, "-")
        lblDetallePartNumber.Text = "Part Number: " & If(Not String.IsNullOrEmpty(item.CurrentPartNumber), item.CurrentPartNumber, "-")
        lblDetallePos.Text = "Posición asignada: " & item.ProposedPosition
        lblDetallePropuesto.Text = "Código Nuevo: " & item.ProposedCode

        ' Extracción segura de la miniatura sin dependencias problemáticas
        Try
            picPreview.Image = ObtenerThumbnailSeguro(item.Document)
        Catch
            picPreview.Image = Nothing
        End Try
    End Sub

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

    ''' <summary>
    ''' Centra la cámara de Inventor y resalta con color celeste vibrante la pieza seleccionada.
    ''' </summary>
    Private Sub HacerZoom3D(item As ElementoPadreItem)
        If item Is Nothing OrElse _invApp Is Nothing Then Return

        Try
            Dim occ As ComponentOccurrence = item.Occurrence
            If occ Is Nothing Then
                occ = BuscarOcurrenciaPorDocumento(_asmDoc.ComponentDefinition.Occurrences, item.DocumentPath)
                item.Occurrence = occ
            End If

            If occ Is Nothing Then Return

            ' Resaltado interactivo en celeste
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
                HabilitarVentanaInventor()
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Resalta un componente en el visor 3D usando HighlightSets nativo de Inventor.
    ''' </summary>
    Private Sub ResaltarElemento(occ As ComponentOccurrence)
        If occ Is Nothing OrElse _invApp Is Nothing Then
            LimpiarResaltado()
            Return
        End If

        Try
            Dim doc As Document = _invApp.ActiveDocument
            If doc Is Nothing Then Return

            LimpiarResaltado()

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
                Try : _highlightSet = doc.CreateHighlightSet() : Catch : End Try
            End If

            If _highlightSet IsNot Nothing Then
                Dim colorCeleste As Inventor.Color = _invApp.TransientObjects.CreateColor(0, 195, 255)
                Try : colorCeleste.Opacity = 0.85 : Catch : End Try
                _highlightSet.Color = colorCeleste

                Try
                    _highlightSet.AddItem(occ)
                Catch
                End Try
            End If

            Try : doc.SelectSet.Clear() : Catch : End Try

            If _invApp.ActiveView IsNot Nothing Then
                _invApp.ActiveView.Update()
            End If

            HabilitarVentanaInventor()
        Catch
        End Try
    End Sub

    Private Sub LimpiarResaltado()
        Try
            If _highlightSet IsNot Nothing Then
                _highlightSet.Clear()
                _highlightSet = Nothing
                If _invApp IsNot Nothing AndAlso _invApp.ActiveView IsNot Nothing Then
                    _invApp.ActiveView.Update()
                End If
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Abre el archivo del componente seleccionado directamente en Inventor.
    ''' </summary>
    Private Sub AbrirElemento(item As ElementoPadreItem)
        If item Is Nothing OrElse String.IsNullOrEmpty(item.DocumentPath) Then Return

        If Not IOFile.Exists(item.DocumentPath) Then
            MessageBox.Show("El archivo no existe en el disco:" & vbCrLf & item.DocumentPath, "Archivo no encontrado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        HabilitarVentanaInventor()

        Try
            _invApp.Documents.Open(item.DocumentPath, True)
            Try
                Dim hwnd As New IntPtr(_invApp.MainFrameHWND)
                If hwnd <> IntPtr.Zero Then SetForegroundWindow(hwnd)
            Catch
            End Try
        Catch ex As Exception
            Try
                System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo() With {
                    .FileName = item.DocumentPath,
                    .UseShellExecute = True
                })
            Catch
            End Try
        End Try

        HabilitarVentanaInventor()
    End Sub

    ' =====================================================================
    ' EVENTOS DE INTERFAZ
    ' =====================================================================

    Private Sub Grid_SelectionChanged(sender As Object, e As EventArgs)
        If _isSyncingSelection Then Return

        _isSyncingSelection = True
        Try
            If _grid.CurrentRow IsNot Nothing Then
                Dim item As ElementoPadreItem = TryCast(_grid.CurrentRow.Tag, ElementoPadreItem)
                ActualizarDetallesItem(item)

                If item IsNot Nothing Then
                    ' Sincronizar selección en el árbol
                    Dim nodoEncontrado = BuscarNodoPorItem(treeEstructura.Nodes, item)
                    If nodoEncontrado IsNot Nothing AndAlso treeEstructura.SelectedNode IsNot nodoEncontrado Then
                        treeEstructura.SelectedNode = nodoEncontrado
                    End If

                    If item.Occurrence IsNot Nothing Then
                        ResaltarElemento(item.Occurrence)
                    End If
                End If
            End If
            HabilitarVentanaInventor()
        Finally
            _isSyncingSelection = False
        End Try
    End Sub

    Private Function BuscarNodoPorItem(nodos As TreeNodeCollection, item As ElementoPadreItem) As TreeNode
        For Each n As TreeNode In nodos
            If n.Tag Is item Then Return n
            Dim subFound = BuscarNodoPorItem(n.Nodes, item)
            If subFound IsNot Nothing Then Return subFound
        Next
        Return Nothing
    End Function

    Private Sub Grid_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        If e.ColumnIndex <> _grid.Columns("ColInc").Index Then
            Dim item As ElementoPadreItem = TryCast(_grid.Rows(e.RowIndex).Tag, ElementoPadreItem)
            If item IsNot Nothing Then
                HacerZoom3D(item)
            End If
        End If
    End Sub

    Private Sub Grid_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs)
        If _grid.IsCurrentCellDirty Then
            _grid.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub Grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return

        Dim item As ElementoPadreItem = TryCast(_grid.Rows(e.RowIndex).Tag, ElementoPadreItem)
        If item Is Nothing Then Return

        Dim colName As String = _grid.Columns(e.ColumnIndex).Name

        If colName = "ColInc" Then
            Dim val = _grid.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            item.Included = If(val IsNot Nothing AndAlso TypeOf val Is Boolean, CBool(val), False)
            AutoNumerarElementos()
        ElseIf colName = "ColPos" Then
            Dim val As String = If(_grid.Rows(e.RowIndex).Cells(e.ColumnIndex).Value IsNot Nothing, _grid.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString(), "")
            item.ProposedPosition = val
            Dim separador As String = If(cmbSeparador.SelectedItem IsNot Nothing, cmbSeparador.SelectedItem.ToString(), ".")
            item.ProposedCode = item.SuggestedPadre & separador & val
            item.ManualOverride = True
            _grid.Rows(e.RowIndex).Cells("ColNuevoCodigo").Value = item.ProposedCode
            ActualizarDetallesItem(item)
        ElseIf colName = "ColNuevoCodigo" Then
            Dim val As String = If(_grid.Rows(e.RowIndex).Cells(e.ColumnIndex).Value IsNot Nothing, _grid.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString(), "")
            item.ProposedCode = val
            item.ManualOverride = True
            ActualizarDetallesItem(item)
        End If
    End Sub

    Private Sub BtnZoom_Click(sender As Object, e As EventArgs)
        If _grid.CurrentRow IsNot Nothing Then
            Dim item As ElementoPadreItem = TryCast(_grid.CurrentRow.Tag, ElementoPadreItem)
            If item IsNot Nothing Then HacerZoom3D(item)
        End If
    End Sub

    Private Sub BtnAbrir_Click(sender As Object, e As EventArgs)
        If _grid.CurrentRow IsNot Nothing Then
            Dim item As ElementoPadreItem = TryCast(_grid.CurrentRow.Tag, ElementoPadreItem)
            If item IsNot Nothing Then AbrirElemento(item)
        End If
    End Sub

    Private Sub CmbEsquema_SelectedIndexChanged(sender As Object, e As EventArgs)
        Select Case cmbEsquema.SelectedIndex
            Case 0 : cmbSeparador.SelectedItem = "."
            Case 1 : cmbSeparador.SelectedItem = "-"
            Case 2 : cmbSeparador.SelectedItem = "_"
        End Select
        AutoNumerarElementos()
    End Sub

    Private Sub BtnMarcarTodos_Click(sender As Object, e As EventArgs)
        For Each item In _items
            item.Included = True
        Next
        AutoNumerarElementos()
    End Sub

    Private Sub BtnDesmarcarTodos_Click(sender As Object, e As EventArgs)
        For Each item In _items
            item.Included = False
        Next
        AutoNumerarElementos()
    End Sub

    Private Sub BtnAutoNumerar_Click(sender As Object, e As EventArgs)
        For Each it In _items
            it.ManualOverride = False
        Next
        AutoNumerarElementos()
        ConstruirArbolJerarquico()
    End Sub

    Private Sub TxtBuscar_TextChanged(sender As Object, e As EventArgs)
        RefrescarGrilla()
    End Sub

    ''' <summary>
    ''' Guarda y aplica los nuevos códigos de ensamble padre y posiciones calculadas
    ''' en los archivos y metadatos de Autodesk Inventor mediante una transacción atómica.
    ''' </summary>
    Private Sub BtnAplicar_Click(sender As Object, e As EventArgs)
        Dim itemsToApply = _items.FindAll(Function(i) i.Included AndAlso Not String.IsNullOrEmpty(i.ProposedCode) AndAlso i.ProposedPosition <> "-")

        If itemsToApply.Count = 0 Then
            MessageBox.Show("No hay elementos seleccionados para numerar.", "Asignar Padres", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim confirmMsg As String = String.Format("Se actualizarán {0} componentes con los nuevos códigos y posiciones." & vbCrLf & vbCrLf &
                                                 "• Nº Almacenamiento (Stock Number): {1}" & vbCrLf &
                                                 "• Nº de Pieza (Part Number): {2}" & vbCrLf &
                                                 "• Propiedad PADRE: {3}" & vbCrLf &
                                                 "• Propiedad POS: {4}" & vbCrLf & vbCrLf &
                                                 "¿Desea aplicar y guardar los cambios ahora?",
                                                 itemsToApply.Count,
                                                 If(chkActualizarStockNumber.Checked, "SÍ", "NO"),
                                                 If(chkActualizarPartNumber.Checked, "SÍ", "NO"),
                                                 If(chkActualizarPadre.Checked, "SÍ", "NO"),
                                                 If(chkActualizarPos.Checked, "SÍ", "NO"))

        If MessageBox.Show(confirmMsg, "Confirmar Numeración de Piezas", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Return
        End If

        Dim modCount As Integer = 0
        Dim oTrans As Transaction = Nothing

        Try
            oTrans = _invApp.TransactionManager.StartTransaction(DirectCast(_asmDoc, Document), "ASIGNAR PADRES - Numeración de Piezas")

            For Each item In itemsToApply
                If item.Document Is Nothing Then Continue For

                Dim doc As Document = item.Document
                Dim changed As Boolean = False

                ' 1. Sincronizar Nº de almacenamiento (Stock Number)
                If chkActualizarStockNumber.Checked AndAlso Not String.IsNullOrEmpty(item.ProposedCode) Then
                    Try
                        Dim dtProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                        dtProps.Item("Stock Number").Value = item.ProposedCode
                        changed = True
                    Catch
                    End Try
                End If

                ' 2. Sincronizar Nº de pieza (Part Number)
                If chkActualizarPartNumber.Checked AndAlso Not String.IsNullOrEmpty(item.ProposedCode) Then
                    Try
                        Dim dtProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                        dtProps.Item("Part Number").Value = item.ProposedCode
                        changed = True
                    Catch
                    End Try
                End If

                ' 3. iProperty personalizada PADRE
                If chkActualizarPadre.Checked AndAlso Not String.IsNullOrEmpty(item.SuggestedPadre) Then
                    SetCustomProp(doc, "PADRE", item.SuggestedPadre)
                    changed = True
                End If

                ' 4. iProperty personalizada POS
                If chkActualizarPos.Checked AndAlso Not String.IsNullOrEmpty(item.ProposedPosition) Then
                    SetCustomProp(doc, "POS", item.ProposedPosition)
                    SetCustomProp(doc, "POSICION", item.ProposedPosition)
                    changed = True
                End If

                If changed Then
                    modCount += 1
                    Try : doc.Update() : Catch : End Try
                End If
            Next

            oTrans.End()
            LimpiarResaltado()

            MessageBox.Show(String.Format("¡Numeración completada con éxito!" & vbCrLf & vbCrLf &
                                          "Archivos actualizados: {0}", modCount),
                            "ASIGNAR PADRES", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            If oTrans IsNot Nothing Then
                Try : oTrans.Abort() : Catch : End Try
            End If
            MessageBox.Show("Error durante la aplicación de códigos: " & ex.Message, "Error en Asignar Padres", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' =====================================================================
    ' FUNCIONES AUXILIARES DE IPROPERTIES DE INVENTOR
    ' =====================================================================

    Private Function ObtenerCodigoBaseEnsamble(doc As Document) As String
        If doc Is Nothing Then Return ""
        Dim stock As String = ObtenerStockNumber(doc)
        If Not String.IsNullOrEmpty(stock) Then Return stock

        Dim pn As String = ObtenerPartNumber(doc)
        If Not String.IsNullOrEmpty(pn) Then Return pn

        Try
            Return IOPath.GetFileNameWithoutExtension(doc.FullFileName)
        Catch
            Return "PADRE"
        End Try
    End Function

    Private Function ObtenerStockNumber(doc As Document) As String
        If doc Is Nothing Then Return ""
        Try
            Dim dtProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
            Return CStr(dtProps.Item("Stock Number").Value)
        Catch
            Return ""
        End Try
    End Function

    Private Function ObtenerPartNumber(doc As Document) As String
        If doc Is Nothing Then Return ""
        Try
            Dim dtProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
            Return CStr(dtProps.Item("Part Number").Value)
        Catch
            Return ""
        End Try
    End Function

    Private Function ObtenerCustomProp(doc As Document, propName As String) As String
        If doc Is Nothing OrElse String.IsNullOrEmpty(propName) Then Return ""
        Try
            Dim customProps As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
            For Each p As Inventor.Property In customProps
                If String.Equals(p.Name.Trim(), propName.Trim(), StringComparison.OrdinalIgnoreCase) Then
                    Return If(p.Value IsNot Nothing, p.Value.ToString().Trim(), "")
                End If
            Next
            Return ""
        Catch
            Return ""
        End Try
    End Function

    Private Sub SetCustomProp(doc As Document, propName As String, propValue As String)
        If doc Is Nothing Then Return
        Try
            Dim customProps As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
            Try
                customProps.Item(propName).Value = propValue
            Catch
                customProps.Add(propValue, propName)
            End Try
        Catch
        End Try
    End Sub
End Class
