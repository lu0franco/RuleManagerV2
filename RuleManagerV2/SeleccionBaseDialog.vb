Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Windows.Forms
Imports Inventor

''' <summary>
''' Diálogo interactivo para Selección Base (IPT).
''' Reemplaza los InputListBox sucesivos de la regla iLogic por una interfaz moderna
''' que permite configurar el tipo de pieza, espesores, dimensiones, materiales,
''' inyectar nuevos espesores de chapa, nuevos diámetros/cuadrados y registrar nuevos materiales.
''' </summary>
Public Class SeleccionBaseDialog
    Inherits Form

    Private ReadOnly _invApp As Inventor.Application
    Private ReadOnly _engine As SeleccionBaseEngine
    Private ReadOnly _doc As Document

    ' Controles principales
    Private _cmbTipoPieza As ComboBox
    Private _chkGuardarPrecargar As CheckBox
    Private _pnlChapa As Panel
    Private _rdoMilimetros As RadioButton
    Private _rdoPulgadas As RadioButton
    Private _cmbEspesorChapa As ComboBox
    Private _btnNuevoEspesor As Button

    Private _pnlMacizo As Panel
    Private _cmbSubTipoFabricacion As ComboBox
    Private _cmbDimensionMacizo As ComboBox
    Private _btnNuevaDimension As Button
    Private _lblDimensionTitulo As Label

    Private _pnlMaterial As Panel
    Private _cmbMaterial As ComboBox
    Private _btnNuevoMaterial As Button
    Private _btnRefrescarMaterial As Button
    Private _btnAbrirLibreria As Button
    Private _lblLibreriaMaterial As Label
    Private _rutaLibreriaManualActual As String = ""

    ' Etiquetas de vista previa
    Private _grpPreview As GroupBox
    Private _lblMaterialBase As Label
    Private _lblMaterialPlano As Label
    Private _lblSubtipoDoc As Label
    Private _lblGeometriaInfo As Label

    Private _btnAceptar As Button
    Private _btnCancelar As Button

    ' =========================================================
    ' PROPIEDADES PÚBLICAS PARA RECUPERAR LA SELECCIÓN
    ' =========================================================
    Public ReadOnly Property TipoPiezaSeleccionado As String
        Get
            Return If(_cmbTipoPieza.SelectedItem IsNot Nothing, _cmbTipoPieza.SelectedItem.ToString(), "CHAPA")
        End Get
    End Property

    Public ReadOnly Property SubTipoFabricacionSeleccionado As String
        Get
            Return If(_cmbSubTipoFabricacion.SelectedItem IsNot Nothing, _cmbSubTipoFabricacion.SelectedItem.ToString(), "TREFILADO")
        End Get
    End Property

    Public ReadOnly Property DimensionOEstiloSeleccionado As String
        Get
            Dim tipo As String = TipoPiezaSeleccionado
            If tipo = "CHAPA" OrElse tipo = "PLANCHA" Then
                Return If(_cmbEspesorChapa.SelectedItem IsNot Nothing, _cmbEspesorChapa.SelectedItem.ToString(), "")
            Else
                Return If(_cmbDimensionMacizo.SelectedItem IsNot Nothing, _cmbDimensionMacizo.SelectedItem.ToString(), "")
            End If
        End Get
    End Property

    Public ReadOnly Property MaterialSeleccionado As String
        Get
            Return If(_cmbMaterial.SelectedItem IsNot Nothing, _cmbMaterial.SelectedItem.ToString(), "")
        End Get
    End Property

    Public ReadOnly Property UnidadesMilimetros As Boolean
        Get
            Return _rdoMilimetros.Checked
        End Get
    End Property

    Public Sub New(invApp As Inventor.Application, engine As SeleccionBaseEngine, doc As Document)
        _invApp = invApp
        _engine = engine
        _doc = doc

        InitializeComponent()
        CargarDatosIniciales()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "RuleManager V2 - Selección Base de Pieza"
        Me.Size = New Size(540, 560)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
        Me.BackColor = System.Drawing.Color.FromArgb(248, 249, 250)

        AddHandler Me.Load, AddressOf OnDialogLoad
        AddHandler Me.Shown, AddressOf OnDialogShown

        ' Header Banner
        Dim pnlHeader As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 60,
            .BackColor = System.Drawing.Color.FromArgb(15, 23, 42)
        }
        Dim lblTitle As New Label With {
            .Text = "Configurador de Tipo y Material de Pieza",
            .ForeColor = System.Drawing.Color.White,
            .Font = New Font("Segoe UI", 11.0F, FontStyle.Bold),
            .Location = New System.Drawing.Point(16, 10),
            .AutoSize = True
        }
        Dim lblSubTitle As New Label With {
            .Text = "Genera iProperties MATERIAL_BASE y MATERIAL_PLANO de forma normalizada",
            .ForeColor = System.Drawing.Color.FromArgb(148, 163, 184),
            .Font = New Font("Segoe UI", 8.5F),
            .Location = New System.Drawing.Point(16, 32),
            .AutoSize = True
        }
        pnlHeader.Controls.AddRange({lblTitle, lblSubTitle})

        Dim curY As Integer = 70

        ' 1. Selector Tipo de Pieza y Casilla de Guardar/Precargar
        Dim lblTipo As New Label With {
            .Text = "1. Tipo de Componente:",
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .Location = New System.Drawing.Point(20, curY),
            .AutoSize = True
        }

        _chkGuardarPrecargar = New CheckBox With {
            .Text = "Guardar/Precargar datos",
            .Location = New System.Drawing.Point(310, curY - 2),
            .AutoSize = True,
            .Checked = AddinSettings.Current.PrecargarGuardarDatosPieza,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Regular),
            .ForeColor = System.Drawing.Color.FromArgb(15, 23, 42),
            .Cursor = Cursors.Hand
        }
        AddHandler _chkGuardarPrecargar.CheckedChanged, AddressOf OnGuardarPrecargarChanged
        curY += 22

        _cmbTipoPieza = New ComboBox With {
            .Location = New System.Drawing.Point(20, curY),
            .Width = 480,
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Font = New Font("Segoe UI", 9.5F)
        }
        _cmbTipoPieza.Items.AddRange(New Object() {"CHAPA", "PLANCHA", "CUADRADO", "REDONDO", "GUIA", "GENERICO"})
        AddHandler _cmbTipoPieza.SelectedIndexChanged, AddressOf OnTipoPiezaChanged
        curY += 36

        ' 2. Panel Chapa
        _pnlChapa = New Panel With {
            .Location = New System.Drawing.Point(20, curY),
            .Width = 480,
            .Height = 70,
            .BackColor = System.Drawing.Color.White,
            .BorderStyle = BorderStyle.FixedSingle
        }
        Dim lblChapaUnits As New Label With {
            .Text = "Unidades:",
            .Location = New System.Drawing.Point(10, 10),
            .AutoSize = True
        }
        _rdoMilimetros = New RadioButton With {
            .Text = "Milímetros",
            .Location = New System.Drawing.Point(80, 8),
            .Checked = True,
            .AutoSize = True
        }
        _rdoPulgadas = New RadioButton With {
            .Text = "Pulgadas",
            .Location = New System.Drawing.Point(180, 8),
            .AutoSize = True
        }
        AddHandler _rdoMilimetros.CheckedChanged, AddressOf OnUnidadesChapaChanged

        Dim lblChapaEspesor As New Label With {
            .Text = "Regla / Espesor:",
            .Location = New System.Drawing.Point(10, 36),
            .AutoSize = True
        }
        _cmbEspesorChapa = New ComboBox With {
            .Location = New System.Drawing.Point(110, 34),
            .Width = 230,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        AddHandler _cmbEspesorChapa.SelectedIndexChanged, AddressOf ActualizarVistaPrevia

        _btnNuevoEspesor = New Button With {
            .Text = "+ Nuevo Espesor",
            .Location = New System.Drawing.Point(348, 32),
            .Width = 120,
            .Height = 26,
            .BackColor = System.Drawing.Color.FromArgb(241, 245, 249)
        }
        AddHandler _btnNuevoEspesor.Click, AddressOf OnNuevoEspesorClicked

        _pnlChapa.Controls.AddRange({lblChapaUnits, _rdoMilimetros, _rdoPulgadas, lblChapaEspesor, _cmbEspesorChapa, _btnNuevoEspesor})

        ' 3. Panel Macizo (Cuadrado / Redondo)
        _pnlMacizo = New Panel With {
            .Location = New System.Drawing.Point(20, curY),
            .Width = 480,
            .Height = 70,
            .BackColor = System.Drawing.Color.White,
            .BorderStyle = BorderStyle.FixedSingle,
            .Visible = False
        }
        Dim lblFab As New Label With {
            .Text = "Tipo Fabricación:",
            .Location = New System.Drawing.Point(10, 10),
            .AutoSize = True
        }
        _cmbSubTipoFabricacion = New ComboBox With {
            .Location = New System.Drawing.Point(120, 8),
            .Width = 140,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        _cmbSubTipoFabricacion.Items.AddRange(New Object() {"TREFILADO", "LAMINADO", "PLASTICO"})
        AddHandler _cmbSubTipoFabricacion.SelectedIndexChanged, AddressOf OnSubTipoFabricacionChanged

        _lblDimensionTitulo = New Label With {
            .Text = "Diámetro / Dimensión:",
            .Location = New System.Drawing.Point(10, 36),
            .AutoSize = True
        }
        _cmbDimensionMacizo = New ComboBox With {
            .Location = New System.Drawing.Point(140, 34),
            .Width = 190,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        AddHandler _cmbDimensionMacizo.SelectedIndexChanged, AddressOf ActualizarVistaPrevia

        _btnNuevaDimension = New Button With {
            .Text = "+ Nueva Medida",
            .Location = New System.Drawing.Point(338, 32),
            .Width = 130,
            .Height = 26,
            .BackColor = System.Drawing.Color.FromArgb(241, 245, 249)
        }
        AddHandler _btnNuevaDimension.Click, AddressOf OnNuevaDimensionClicked

        _pnlMacizo.Controls.AddRange({
            lblFab, _cmbSubTipoFabricacion, _lblDimensionTitulo, _cmbDimensionMacizo, _btnNuevaDimension
        })
        curY += 76

        ' 4. Panel Material
        _pnlMaterial = New Panel With {
            .Location = New System.Drawing.Point(20, curY),
            .Width = 480,
            .Height = 72,
            .BackColor = System.Drawing.Color.White,
            .BorderStyle = BorderStyle.FixedSingle
        }
        Dim lblMatTitle As New Label With {
            .Text = "Material:",
            .Location = New System.Drawing.Point(10, 8),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .AutoSize = True
        }
        _lblLibreriaMaterial = New Label With {
            .Text = "(Librería: METALES DB)",
            .Location = New System.Drawing.Point(75, 8),
            .ForeColor = System.Drawing.Color.FromArgb(71, 85, 105),
            .AutoSize = True
        }
        _cmbMaterial = New ComboBox With {
            .Location = New System.Drawing.Point(10, 32),
            .Width = 240,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        AddHandler _cmbMaterial.SelectedIndexChanged, AddressOf ActualizarVistaPrevia

        _btnRefrescarMaterial = New Button With {
            .Text = "🔄 Recargar",
            .Location = New System.Drawing.Point(256, 30),
            .Width = 72,
            .Height = 27,
            .BackColor = System.Drawing.Color.FromArgb(241, 245, 249),
            .Font = New Font("Segoe UI", 8.0F)
        }
        AddHandler _btnRefrescarMaterial.Click, AddressOf OnRefrescarMaterialClicked

        _btnNuevoMaterial = New Button With {
            .Text = "+ Nuevo",
            .Location = New System.Drawing.Point(332, 30),
            .Width = 68,
            .Height = 27,
            .BackColor = System.Drawing.Color.FromArgb(241, 245, 249),
            .Font = New Font("Segoe UI", 8.0F)
        }
        AddHandler _btnNuevoMaterial.Click, AddressOf OnNuevoMaterialClicked

        _btnAbrirLibreria = New Button With {
            .Text = "📂 .adsklib",
            .Location = New System.Drawing.Point(404, 30),
            .Width = 66,
            .Height = 27,
            .BackColor = System.Drawing.Color.FromArgb(241, 245, 249),
            .Font = New Font("Segoe UI", 8.0F)
        }
        AddHandler _btnAbrirLibreria.Click, AddressOf OnAbrirLibreriaClicked

        _pnlMaterial.Controls.AddRange({lblMatTitle, _lblLibreriaMaterial, _cmbMaterial, _btnRefrescarMaterial, _btnNuevoMaterial, _btnAbrirLibreria})
        curY += 80

        ' 5. Vista Previa de Propiedades Generadas y Geometría
        _grpPreview = New GroupBox With {
            .Text = "Vista Previa de Propiedades y Geometría",
            .Location = New System.Drawing.Point(20, curY),
            .Width = 480,
            .Height = 132,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        }
        Dim l1 As New Label With {.Text = "MATERIAL_BASE:", .Location = New System.Drawing.Point(12, 22), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F)}
        _lblMaterialBase = New Label With {.Text = "CHAPA", .Location = New System.Drawing.Point(140, 22), .AutoSize = True, .ForeColor = System.Drawing.Color.FromArgb(2, 132, 199)}

        Dim l2 As New Label With {.Text = "MATERIAL_PLANO:", .Location = New System.Drawing.Point(12, 46), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F)}
        _lblMaterialPlano = New Label With {.Text = "=<MATERIAL_BASE> <SHEET METAL RULE> <MATERIAL>", .Location = New System.Drawing.Point(140, 46), .AutoSize = True, .ForeColor = System.Drawing.Color.FromArgb(16, 185, 129)}

        Dim l3 As New Label With {.Text = "SubTipo Inventor:", .Location = New System.Drawing.Point(12, 70), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F)}
        _lblSubtipoDoc = New Label With {.Text = "Sheet Metal Document", .Location = New System.Drawing.Point(140, 70), .AutoSize = True, .ForeColor = System.Drawing.Color.FromArgb(100, 116, 139)}

        Dim l4 As New Label With {.Text = "Geometría Boceto:", .Location = New System.Drawing.Point(12, 94), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F)}
        _lblGeometriaInfo = New Label With {.Text = "Sin boceto previo", .Location = New System.Drawing.Point(140, 94), .AutoSize = True, .ForeColor = System.Drawing.Color.FromArgb(168, 85, 247)}

        _grpPreview.Controls.AddRange({l1, _lblMaterialBase, l2, _lblMaterialPlano, l3, _lblSubtipoDoc, l4, _lblGeometriaInfo})
        curY += 140

        ' Botones inferiores
        _btnAceptar = New Button With {
            .Text = "Aplicar a Pieza",
            .Location = New System.Drawing.Point(270, curY),
            .Width = 130,
            .Height = 32,
            .BackColor = System.Drawing.Color.FromArgb(14, 165, 233),
            .ForeColor = System.Drawing.Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
        }
        AddHandler _btnAceptar.Click, AddressOf OnAceptarClicked

        _btnCancelar = New Button With {
            .Text = "Cancelar",
            .Location = New System.Drawing.Point(410, curY),
            .Width = 90,
            .Height = 32,
            .DialogResult = DialogResult.Cancel
        }

        Me.Controls.AddRange({pnlHeader, _cmbTipoPieza, lblTipo, _chkGuardarPrecargar, _pnlChapa, _pnlMacizo, _pnlMaterial, _grpPreview, _btnAceptar, _btnCancelar})
    End Sub

    Private Sub OnDialogLoad(sender As Object, e As EventArgs)
        ReacomodarLayout()
        ActualizarVistaPrevia()
    End Sub

    Private Sub OnDialogShown(sender As Object, e As EventArgs)
        ReacomodarLayout()
        ActualizarVistaPrevia()
        Me.Refresh()
    End Sub

    Private Function DetectarTipoInicial() As String
        Try
            If _doc IsNot Nothing Then
                ' 1. Verificar si ya existe la iProperty MATERIAL_BASE
                Dim customProps As PropertySet = Nothing
                Try
                    customProps = _doc.PropertySets.Item("Inventor User Defined Properties")
                Catch
                End Try

                If customProps IsNot Nothing Then
                    For Each prop As Inventor.Property In customProps
                        If prop.Name.Equals("MATERIAL_BASE", StringComparison.OrdinalIgnoreCase) Then
                            Dim val As String = If(prop.Value IsNot Nothing, prop.Value.ToString().Trim().ToUpper(), "")
                            If val.StartsWith("CHAPA") Then Return "CHAPA"
                            If val.StartsWith("PLANCHA") Then Return "PLANCHA"
                            If val.StartsWith("CUADRADO") Then Return "CUADRADO"
                            If val.StartsWith("REDONDO") Then Return "REDONDO"
                            If val.StartsWith("GUIA") Then Return "GUIA"
                            If val.StartsWith("GENERICO") Then Return "GENERICO"
                        End If
                    Next
                End If

                ' 2. Si no tiene iProperty, verificar subtipo nativo de Inventor
                If _doc.SubType = SeleccionBaseEngine.SUBTYPE_SHEET_METAL Then
                    Return "CHAPA"
                End If
            End If
        Catch
        End Try

        ' 3. Si la opción de precargar está activa, recordar el último tipo utilizado
        If _chkGuardarPrecargar IsNot Nothing AndAlso _chkGuardarPrecargar.Checked Then
            If Not String.IsNullOrEmpty(AddinSettings.Current.UltimoTipoPieza) Then
                Return AddinSettings.Current.UltimoTipoPieza
            End If
        ElseIf AddinSettings.Current.PrecargarGuardarDatosPieza Then
            If Not String.IsNullOrEmpty(AddinSettings.Current.UltimoTipoPieza) Then
                Return AddinSettings.Current.UltimoTipoPieza
            End If
        End If

        Return "CHAPA"
    End Function

    Private Sub SincronizarValoresDesdeDocumento()
        Try
            If _doc Is Nothing Then Exit Sub

            ' 1. Sincronizar material activo de la pieza si está presente en la librería activa
            Dim partDoc As PartDocument = TryCast(_doc, PartDocument)
            If partDoc IsNot Nothing AndAlso partDoc.ComponentDefinition.Material IsNot Nothing Then
                Dim docMatName As String = partDoc.ComponentDefinition.Material.Name
                If _cmbMaterial.Items.Contains(docMatName) Then
                    _cmbMaterial.SelectedItem = docMatName
                End If
            End If

            ' 2. Sincronizar subtipo de fabricación o medida desde MATERIAL_BASE si existe
            Dim customProps As PropertySet = Nothing
            Try
                customProps = _doc.PropertySets.Item("Inventor User Defined Properties")
            Catch
            End Try

            If customProps IsNot Nothing Then
                For Each prop As Inventor.Property In customProps
                    If prop.Name.Equals("MATERIAL_BASE", StringComparison.OrdinalIgnoreCase) Then
                        Dim val As String = If(prop.Value IsNot Nothing, prop.Value.ToString().Trim(), "")
                        Dim partes As String() = val.Split(" "c)
                        If partes.Length >= 2 Then
                            Dim subFabCandidate As String = partes(1).ToUpper()
                            If _cmbSubTipoFabricacion.Items.Contains(subFabCandidate) Then
                                _cmbSubTipoFabricacion.SelectedItem = subFabCandidate
                            End If
                        End If
                        If partes.Length >= 3 Then
                            Dim dimCandidate As String = val.Substring(val.IndexOf(partes(1)) + partes(1).Length).Trim()
                            If _pnlMacizo.Visible Then
                                For Each item In _cmbDimensionMacizo.Items
                                    If item.ToString().StartsWith(dimCandidate, StringComparison.OrdinalIgnoreCase) OrElse
                                       dimCandidate.StartsWith(item.ToString(), StringComparison.OrdinalIgnoreCase) Then
                                        _cmbDimensionMacizo.SelectedItem = item
                                        Exit For
                                    End If
                                Next
                            ElseIf _pnlChapa.Visible Then
                                For Each item In _cmbEspesorChapa.Items
                                    If item.ToString().Equals(dimCandidate, StringComparison.OrdinalIgnoreCase) Then
                                        _cmbEspesorChapa.SelectedItem = item
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    End If
                Next
            End If
        Catch
        End Try
    End Sub

    Private Sub PrecargarDatosPorTipo(tipo As String)
        If _chkGuardarPrecargar Is Nothing OrElse Not _chkGuardarPrecargar.Checked Then Exit Sub

        Dim s = AddinSettings.Current

        Select Case tipo
            Case "CHAPA"
                If s.UltimaChapaUnidadesMm Then
                    _rdoMilimetros.Checked = True
                Else
                    _rdoPulgadas.Checked = True
                End If
                CargarEspesoresChapa()
                If Not String.IsNullOrEmpty(s.UltimaChapaEspesor) AndAlso _cmbEspesorChapa.Items.Contains(s.UltimaChapaEspesor) Then
                    _cmbEspesorChapa.SelectedItem = s.UltimaChapaEspesor
                End If
                If Not String.IsNullOrEmpty(s.UltimaChapaMaterial) AndAlso _cmbMaterial.Items.Contains(s.UltimaChapaMaterial) Then
                    _cmbMaterial.SelectedItem = s.UltimaChapaMaterial
                End If

            Case "PLANCHA"
                CargarEspesoresPlancha()
                If Not String.IsNullOrEmpty(s.UltimaPlanchaEspesor) AndAlso _cmbEspesorChapa.Items.Contains(s.UltimaPlanchaEspesor) Then
                    _cmbEspesorChapa.SelectedItem = s.UltimaPlanchaEspesor
                End If
                If Not String.IsNullOrEmpty(s.UltimaPlanchaMaterial) AndAlso _cmbMaterial.Items.Contains(s.UltimaPlanchaMaterial) Then
                    _cmbMaterial.SelectedItem = s.UltimaPlanchaMaterial
                End If

            Case "CUADRADO"
                If Not String.IsNullOrEmpty(s.UltimoCuadradoSubTipo) AndAlso _cmbSubTipoFabricacion.Items.Contains(s.UltimoCuadradoSubTipo) Then
                    _cmbSubTipoFabricacion.SelectedItem = s.UltimoCuadradoSubTipo
                End If
                CargarDimensionesCuadrado()
                If Not String.IsNullOrEmpty(s.UltimoCuadradoDimension) AndAlso _cmbDimensionMacizo.Items.Contains(s.UltimoCuadradoDimension) Then
                    _cmbDimensionMacizo.SelectedItem = s.UltimoCuadradoDimension
                End If
                CargarMateriales(ObtenerNombreLibreriaActiva())
                If Not String.IsNullOrEmpty(s.UltimoCuadradoMaterial) AndAlso _cmbMaterial.Items.Contains(s.UltimoCuadradoMaterial) Then
                    _cmbMaterial.SelectedItem = s.UltimoCuadradoMaterial
                End If

            Case "REDONDO"
                If Not String.IsNullOrEmpty(s.UltimoRedondoSubTipo) AndAlso _cmbSubTipoFabricacion.Items.Contains(s.UltimoRedondoSubTipo) Then
                    _cmbSubTipoFabricacion.SelectedItem = s.UltimoRedondoSubTipo
                End If
                CargarDiametrosRedondo()
                If Not String.IsNullOrEmpty(s.UltimoRedondoDimension) AndAlso _cmbDimensionMacizo.Items.Contains(s.UltimoRedondoDimension) Then
                    _cmbDimensionMacizo.SelectedItem = s.UltimoRedondoDimension
                End If
                CargarMateriales(ObtenerNombreLibreriaActiva())
                If Not String.IsNullOrEmpty(s.UltimoRedondoMaterial) AndAlso _cmbMaterial.Items.Contains(s.UltimoRedondoMaterial) Then
                    _cmbMaterial.SelectedItem = s.UltimoRedondoMaterial
                End If

            Case "GENERICO"
                If Not String.IsNullOrEmpty(s.UltimoGenericoMaterial) AndAlso _cmbMaterial.Items.Contains(s.UltimoGenericoMaterial) Then
                    _cmbMaterial.SelectedItem = s.UltimoGenericoMaterial
                End If

            Case "GUIA"
                If Not String.IsNullOrEmpty(s.UltimaGuiaMaterial) AndAlso _cmbMaterial.Items.Contains(s.UltimaGuiaMaterial) Then
                    _cmbMaterial.SelectedItem = s.UltimaGuiaMaterial
                End If
        End Select
    End Sub

    Private Sub OnGuardarPrecargarChanged(sender As Object, e As EventArgs)
        AddinSettings.Current.PrecargarGuardarDatosPieza = _chkGuardarPrecargar.Checked
        AddinSettings.Save()

        If _chkGuardarPrecargar.Checked AndAlso _cmbTipoPieza.SelectedItem IsNot Nothing Then
            PrecargarDatosPorTipo(_cmbTipoPieza.SelectedItem.ToString())
            ActualizarVistaPrevia()
        End If
    End Sub

    Private Sub CargarDatosIniciales()
        ' 1. Asegurar subtipo de fabricación por defecto
        If _cmbSubTipoFabricacion.Items.Count > 0 AndAlso _cmbSubTipoFabricacion.SelectedIndex < 0 Then
            _cmbSubTipoFabricacion.SelectedIndex = 0
        End If

        ' 2. Detectar tipo existente en la pieza o predeterminado
        Dim tipoInicial As String = DetectarTipoInicial()
        Dim idx As Integer = _cmbTipoPieza.FindStringExact(tipoInicial)
        If idx < 0 Then idx = 0

        ' 3. Sincronizar el ComboBox de tipo de pieza
        _cmbTipoPieza.SelectedIndex = idx

        ' 4. Forzar explícitamente la ejecución de configuración completa
        EjecutarCambioTipoPieza(tipoInicial)

        ' 5. Sincronizar valores específicos del documento
        SincronizarValoresDesdeDocumento()

        ' 6. Reacomodar y actualizar
        ReacomodarLayout()
        ActualizarVistaPrevia()
    End Sub

    Private Sub ReacomodarLayout()
        Dim topPaneles As Integer = _cmbTipoPieza.Bottom + 10
        _pnlChapa.Top = topPaneles
        _pnlMacizo.Top = topPaneles

        Dim topMaterial As Integer
        If _pnlMacizo.Visible Then
            topMaterial = _pnlMacizo.Bottom + 10
        ElseIf _pnlChapa.Visible Then
            topMaterial = _pnlChapa.Bottom + 10
        Else
            topMaterial = _cmbTipoPieza.Bottom + 10
        End If

        _pnlMaterial.Top = topMaterial
        _grpPreview.Top = If(_pnlMaterial.Visible, _pnlMaterial.Bottom + 10, topMaterial)
        _btnAceptar.Top = _grpPreview.Bottom + 12
        _btnCancelar.Top = _grpPreview.Bottom + 12
        Me.ClientSize = New Size(520, _btnAceptar.Bottom + 16)
    End Sub

    Private Sub OnTipoPiezaChanged(sender As Object, e As EventArgs)
        If _cmbTipoPieza.SelectedItem Is Nothing Then Exit Sub
        Dim tipo As String = _cmbTipoPieza.SelectedItem.ToString()
        EjecutarCambioTipoPieza(tipo)
    End Sub

    Private Sub EjecutarCambioTipoPieza(tipo As String)
        Select Case tipo
            Case "CHAPA"
                _pnlChapa.Visible = True
                _pnlMacizo.Visible = False
                _pnlMaterial.Visible = True
                _lblLibreriaMaterial.Text = "(Librería activa: METALES DB)"
                CargarEspesoresChapa()
                CargarMateriales("METALES DB")
                If _chkGuardarPrecargar IsNot Nothing AndAlso _chkGuardarPrecargar.Checked Then
                    PrecargarDatosPorTipo("CHAPA")
                End If

            Case "PLANCHA"
                _pnlChapa.Visible = True
                _pnlMacizo.Visible = False
                _pnlMaterial.Visible = True
                _lblLibreriaMaterial.Text = "(Librería activa: PLASTICOS DB)"
                CargarEspesoresPlancha()
                CargarMateriales("PLASTICOS DB")
                If _chkGuardarPrecargar IsNot Nothing AndAlso _chkGuardarPrecargar.Checked Then
                    PrecargarDatosPorTipo("PLANCHA")
                End If

            Case "CUADRADO"
                _pnlChapa.Visible = False
                _pnlMacizo.Visible = True
                _pnlMaterial.Visible = True
                _lblDimensionTitulo.Text = "Lado Cuadrado:"
                If _cmbSubTipoFabricacion.SelectedIndex < 0 Then _cmbSubTipoFabricacion.SelectedIndex = 0
                CargarDimensionesCuadrado()
                CargarMateriales(ObtenerNombreLibreriaActiva())
                If _chkGuardarPrecargar IsNot Nothing AndAlso _chkGuardarPrecargar.Checked Then
                    PrecargarDatosPorTipo("CUADRADO")
                End If

            Case "REDONDO"
                _pnlChapa.Visible = False
                _pnlMacizo.Visible = True
                _pnlMaterial.Visible = True
                _lblDimensionTitulo.Text = "Diámetro Redondo:"
                If _cmbSubTipoFabricacion.SelectedIndex < 0 Then _cmbSubTipoFabricacion.SelectedIndex = 0
                CargarDiametrosRedondo()
                CargarMateriales(ObtenerNombreLibreriaActiva())
                If _chkGuardarPrecargar IsNot Nothing AndAlso _chkGuardarPrecargar.Checked Then
                    PrecargarDatosPorTipo("REDONDO")
                End If

            Case "GENERICO"
                _pnlChapa.Visible = False
                _pnlMacizo.Visible = False
                _pnlMaterial.Visible = False
                If _chkGuardarPrecargar IsNot Nothing AndAlso _chkGuardarPrecargar.Checked Then
                    PrecargarDatosPorTipo("GENERICO")
                End If

            Case "GUIA"
                _pnlChapa.Visible = False
                _pnlMacizo.Visible = False
                _pnlMaterial.Visible = False
                If _chkGuardarPrecargar IsNot Nothing AndAlso _chkGuardarPrecargar.Checked Then
                    PrecargarDatosPorTipo("GUIA")
                End If
        End Select

        ReacomodarLayout()
        ActualizarVistaPrevia()
    End Sub

    Private Sub OnUnidadesChapaChanged(sender As Object, e As EventArgs)
        If _cmbTipoPieza.SelectedItem.ToString() = "CHAPA" Then
            CargarEspesoresChapa()
        End If
    End Sub

    Private Function ObtenerNombreLibreriaActiva() As String
        Dim tipo As String = If(_cmbTipoPieza.SelectedItem IsNot Nothing, _cmbTipoPieza.SelectedItem.ToString(), "CHAPA")
        Dim esPlastico As Boolean = (tipo = "PLANCHA") OrElse
            ((tipo = "REDONDO" OrElse tipo = "CUADRADO") AndAlso _cmbSubTipoFabricacion.SelectedItem IsNot Nothing AndAlso _cmbSubTipoFabricacion.SelectedItem.ToString() = "PLASTICO")
        Return If(esPlastico, "PLASTICOS DB", "METALES DB")
    End Function

    Private Sub OnSubTipoFabricacionChanged(sender As Object, e As EventArgs)
        Dim libName As String = ObtenerNombreLibreriaActiva()
        CargarMateriales(libName)
        ActualizarVistaPrevia()
    End Sub

    Private Sub CargarEspesoresChapa()
        _cmbEspesorChapa.Items.Clear()
        Dim lista As List(Of String)
        If _rdoMilimetros.Checked Then
            lista = _engine.ObtenerEspesoresChapaMm()
        Else
            lista = _engine.ObtenerEspesoresChapaPulg()
        End If
        For Each item In lista
            _cmbEspesorChapa.Items.Add(item)
        Next
        If _cmbEspesorChapa.Items.Count > 0 Then _cmbEspesorChapa.SelectedIndex = 0
    End Sub

    Private Sub CargarEspesoresPlancha()
        _cmbEspesorChapa.Items.Clear()
        Dim lista = _engine.ObtenerEspesoresPlancha()
        For Each item In lista
            _cmbEspesorChapa.Items.Add(item)
        Next
        If _cmbEspesorChapa.Items.Count > 0 Then _cmbEspesorChapa.SelectedIndex = 0
    End Sub

    Private Sub CargarDimensionesCuadrado()
        _cmbDimensionMacizo.Items.Clear()
        Dim lista = _engine.ObtenerDimensionesCuadrado()
        For Each item In lista
            _cmbDimensionMacizo.Items.Add(item)
        Next
        If _cmbDimensionMacizo.Items.Count > 0 Then _cmbDimensionMacizo.SelectedIndex = 0
    End Sub

    Private Sub CargarDiametrosRedondo()
        _cmbDimensionMacizo.Items.Clear()
        Dim lista = _engine.ObtenerDiametrosRedondo()
        For Each item In lista
            _cmbDimensionMacizo.Items.Add(item)
        Next
        If _cmbDimensionMacizo.Items.Count > 0 Then _cmbDimensionMacizo.SelectedIndex = 0
    End Sub

    ' =========================================================
    ' CARGAR MATERIALES DIRECTAMENTE DE LA LIBRERÍA DE INVENTOR
    ' =========================================================
    Private Sub CargarMateriales(nombreLibreria As String)
        _cmbMaterial.Items.Clear()
        Dim seleccionPrevia As String = If(_cmbMaterial.SelectedItem IsNot Nothing, _cmbMaterial.SelectedItem.ToString(), "")

        Dim lista = _engine.ObtenerMaterialesDeLibreria(nombreLibreria, _doc, _rutaLibreriaManualActual)
        For Each item In lista
            _cmbMaterial.Items.Add(item)
        Next

        If lista.Count > 0 Then
            _lblLibreriaMaterial.Text = "(Librería: " & nombreLibreria & " • " & lista.Count & " materiales leídos directamente)"
            _lblLibreriaMaterial.ForeColor = System.Drawing.Color.FromArgb(22, 101, 52)

            If Not String.IsNullOrEmpty(seleccionPrevia) AndAlso _cmbMaterial.Items.Contains(seleccionPrevia) Then
                _cmbMaterial.SelectedItem = seleccionPrevia
            Else
                _cmbMaterial.SelectedIndex = 0
            End If
        Else
            _lblLibreriaMaterial.Text = "(⚠️ 0 materiales en '" & nombreLibreria & "' • Use Recargar o Abrir .adsklib)"
            _lblLibreriaMaterial.ForeColor = System.Drawing.Color.FromArgb(220, 38, 38)
        End If
    End Sub

    Private Sub OnRefrescarMaterialClicked(sender As Object, e As EventArgs)
        Dim libName As String = ObtenerNombreLibreriaActiva()
        CargarMateriales(libName)
        MessageBox.Show("Librería '" & libName & "' leída directamente de Autodesk Inventor." & vbCrLf &
                        "Total materiales leídos de la librería: " & _cmbMaterial.Items.Count,
                        "Librería Actualizada", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub OnAbrirLibreriaClicked(sender As Object, e As EventArgs)
        Dim libName As String = ObtenerNombreLibreriaActiva()
        Using dlg As New OpenFileDialog()
            dlg.Title = "Seleccionar archivo de librería Autodesk para " & libName
            dlg.Filter = "Librerías de Material Autodesk (*.adsklib)|*.adsklib|Todos los archivos (*.*)|*.*"
            dlg.InitialDirectory = "D:\Bibliotecas\Materiales"
            If Not IO.Directory.Exists(dlg.InitialDirectory) Then
                dlg.InitialDirectory = "C:\"
            End If

            If dlg.ShowDialog() = DialogResult.OK Then
                _rutaLibreriaManualActual = dlg.FileName
                Dim libAsset = _engine.ObtenerLibreria(libName, _rutaLibreriaManualActual)
                CargarMateriales(libName)
                MessageBox.Show("Librería cargada desde: " & vbCrLf & _rutaLibreriaManualActual & vbCrLf &
                                "Materiales leídos directamente: " & _cmbMaterial.Items.Count,
                                "Librería Cargada", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Using
    End Sub

    ' =========================================================
    ' EVENTOS DE INYECCIÓN DINÁMICA
    ' =========================================================

    Private Sub OnNuevoEspesorClicked(sender As Object, e As EventArgs)
        Dim input As String = Microsoft.VisualBasic.Interaction.InputBox(
            "Ingrese el nuevo espesor de chapa en milímetros (ej: 4.5):",
            "Inyectar Regla de Chapa", "4.5")

        If String.IsNullOrWhiteSpace(input) Then Exit Sub

        Dim num As Double
        If Double.TryParse(input.Replace(","c, "."c), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, num) Then
            Dim nombreRegla As String = num.ToString() & "mm"
            _engine.InyectarNuevoEspesorChapa(_doc, nombreRegla, num)
            CargarEspesoresChapa()
            _cmbEspesorChapa.SelectedItem = nombreRegla
            MessageBox.Show("Regla de chapa '" & nombreRegla & "' creada e inyectada exitosamente.", "Inyección Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Formato numérico no válido.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub OnNuevaDimensionClicked(sender As Object, e As EventArgs)
        Dim tipo As String = _cmbTipoPieza.SelectedItem.ToString()
        Dim prompt As String = If(tipo = "REDONDO", "Ingrese nuevo diámetro (ej: 18mm o 2 1/4""): ", "Ingrese nueva medida de lado cuadrado (ej: 22mm o 1 3/8""): ")
        Dim input As String = Microsoft.VisualBasic.Interaction.InputBox(prompt, "Agregar Medida", "")

        If String.IsNullOrWhiteSpace(input) Then Exit Sub

        If tipo = "REDONDO" Then
            _engine.InyectarNuevoDiametro(input)
            CargarDiametrosRedondo()
            _cmbDimensionMacizo.SelectedItem = input
        Else
            _engine.InyectarNuevaDimensionCuadrado(input)
            CargarDimensionesCuadrado()
            _cmbDimensionMacizo.SelectedItem = input
        End If

        MessageBox.Show("Medida '" & input & "' guardada y registrada para futuros proyectos.", "Guardado", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub OnNuevoMaterialClicked(sender As Object, e As EventArgs)
        Dim libName As String = ObtenerNombreLibreriaActiva()
        Dim esPlastico As Boolean = (libName = "PLASTICOS DB")

        Dim prompt As String = "Ingrese el nombre del nuevo material para guardar físicamente en la librería '" & libName & "':"
        Dim input As String = Microsoft.VisualBasic.Interaction.InputBox(prompt, "Nuevo Material en " & libName, "")

        If String.IsNullOrWhiteSpace(input) Then Exit Sub

        Dim exito As Boolean = _engine.InyectarNuevoMaterial(input.Trim(), esPlastico, _doc)
        If exito Then
            ' Recargar directamente de la librería física
            CargarMateriales(libName)
            If _cmbMaterial.Items.Contains(input.Trim()) Then
                _cmbMaterial.SelectedItem = input.Trim()
            End If
            MessageBox.Show("Material '" & input.Trim() & "' guardado físicamente en la librería " & libName & " (.adsklib) y listo para usar.",
                            "Material Guardado en Librería", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub ActualizarVistaPrevia()
        If _cmbTipoPieza.SelectedItem Is Nothing Then Exit Sub

        Dim tipo As String = _cmbTipoPieza.SelectedItem.ToString()
        Dim subTipo As String = If(_cmbSubTipoFabricacion.SelectedItem IsNot Nothing, _cmbSubTipoFabricacion.SelectedItem.ToString(), "")
        Dim dimEstilo As String = If(tipo = "CHAPA" OrElse tipo = "PLANCHA",
                                    If(_cmbEspesorChapa.SelectedItem IsNot Nothing, _cmbEspesorChapa.SelectedItem.ToString(), ""),
                                    If(_cmbDimensionMacizo.SelectedItem IsNot Nothing, _cmbDimensionMacizo.SelectedItem.ToString(), ""))

        Select Case tipo
            Case "CHAPA"
                _lblMaterialBase.Text = "CHAPA"
                _lblMaterialPlano.Text = "=<MATERIAL_BASE> <SHEET METAL RULE> <MATERIAL>"
                _lblSubtipoDoc.Text = "Sheet Metal Document ({9C464203...})"
                _lblGeometriaInfo.Text = "N/A (Definido por pestaña/operaciones de chapa)"
                _lblGeometriaInfo.ForeColor = System.Drawing.Color.FromArgb(100, 116, 139)

            Case "PLANCHA"
                _lblMaterialBase.Text = "PLANCHA"
                _lblMaterialPlano.Text = "=<MATERIAL_BASE> <SHEET METAL RULE> <MATERIAL>"
                _lblSubtipoDoc.Text = "Sheet Metal Document ({9C464203...})"
                _lblGeometriaInfo.Text = "N/A (Definido por pestaña/operaciones de chapa)"
                _lblGeometriaInfo.ForeColor = System.Drawing.Color.FromArgb(100, 116, 139)

            Case "CUADRADO"
                _lblMaterialBase.Text = "CUADRADO " & subTipo & " " & dimEstilo & " "
                _lblMaterialPlano.Text = "=<MATERIAL_BASE> <MATERIAL>"
                _lblSubtipoDoc.Text = "Standard Part Document ({4D29B490...})"

                Dim medLimpia As String = SeleccionBaseEngine.ExtraerMedidaLimpia(dimEstilo)
                _lblGeometriaInfo.Text = "Boceto / iFeature: Cuadrado Lado " & medLimpia & " (Plano guiado por cursor al aplicar)"
                _lblGeometriaInfo.ForeColor = System.Drawing.Color.FromArgb(16, 185, 129)

            Case "REDONDO"
                _lblMaterialBase.Text = "REDONDO " & subTipo & " " & dimEstilo & " "
                _lblMaterialPlano.Text = "=<MATERIAL_BASE> <MATERIAL>"
                _lblSubtipoDoc.Text = "Standard Part Document ({4D29B490...})"

                Dim medLimpia As String = SeleccionBaseEngine.ExtraerMedidaLimpia(dimEstilo)
                _lblGeometriaInfo.Text = "Boceto / iFeature: Círculo Diámetro " & medLimpia & " (Plano guiado por cursor al aplicar)"
                _lblGeometriaInfo.ForeColor = System.Drawing.Color.FromArgb(16, 185, 129)

            Case "GENERICO"
                _lblMaterialBase.Text = "GENERICO"
                _lblMaterialPlano.Text = "=<MATERIAL_BASE> <MATERIAL>"
                _lblSubtipoDoc.Text = "Standard Part Document"
                _lblGeometriaInfo.Text = "N/A"
                _lblGeometriaInfo.ForeColor = System.Drawing.Color.FromArgb(100, 116, 139)

            Case "GUIA"
                _lblMaterialBase.Text = "GUIA"
                _lblMaterialPlano.Text = "=<MATERIAL_BASE> <MATERIAL>"
                _lblSubtipoDoc.Text = "Standard Part (BOM: Referencia)"
                _lblGeometriaInfo.Text = "N/A"
                _lblGeometriaInfo.ForeColor = System.Drawing.Color.FromArgb(100, 116, 139)
        End Select
    End Sub

    Private Sub OnAceptarClicked(sender As Object, e As EventArgs)
        If _cmbTipoPieza.SelectedItem Is Nothing Then
            MessageBox.Show("Por favor seleccione un tipo de componente.", "Selección Base", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        ' Guardar / Precargar configuración según el estado de la casilla
        Dim s = AddinSettings.Current
        s.PrecargarGuardarDatosPieza = _chkGuardarPrecargar.Checked

        If _chkGuardarPrecargar.Checked Then
            Dim tipo As String = _cmbTipoPieza.SelectedItem.ToString()
            s.UltimoTipoPieza = tipo

            Select Case tipo
                Case "CHAPA"
                    s.UltimaChapaUnidadesMm = _rdoMilimetros.Checked
                    s.UltimaChapaEspesor = If(_cmbEspesorChapa.SelectedItem IsNot Nothing, _cmbEspesorChapa.SelectedItem.ToString(), "")
                    s.UltimaChapaMaterial = If(_cmbMaterial.SelectedItem IsNot Nothing, _cmbMaterial.SelectedItem.ToString(), "")

                Case "PLANCHA"
                    s.UltimaPlanchaEspesor = If(_cmbEspesorChapa.SelectedItem IsNot Nothing, _cmbEspesorChapa.SelectedItem.ToString(), "")
                    s.UltimaPlanchaMaterial = If(_cmbMaterial.SelectedItem IsNot Nothing, _cmbMaterial.SelectedItem.ToString(), "")

                Case "CUADRADO"
                    s.UltimoCuadradoSubTipo = If(_cmbSubTipoFabricacion.SelectedItem IsNot Nothing, _cmbSubTipoFabricacion.SelectedItem.ToString(), "TREFILADO")
                    s.UltimoCuadradoDimension = If(_cmbDimensionMacizo.SelectedItem IsNot Nothing, _cmbDimensionMacizo.SelectedItem.ToString(), "")
                    s.UltimoCuadradoMaterial = If(_cmbMaterial.SelectedItem IsNot Nothing, _cmbMaterial.SelectedItem.ToString(), "")

                Case "REDONDO"
                    s.UltimoRedondoSubTipo = If(_cmbSubTipoFabricacion.SelectedItem IsNot Nothing, _cmbSubTipoFabricacion.SelectedItem.ToString(), "TREFILADO")
                    s.UltimoRedondoDimension = If(_cmbDimensionMacizo.SelectedItem IsNot Nothing, _cmbDimensionMacizo.SelectedItem.ToString(), "")
                    s.UltimoRedondoMaterial = If(_cmbMaterial.SelectedItem IsNot Nothing, _cmbMaterial.SelectedItem.ToString(), "")

                Case "GENERICO"
                    s.UltimoGenericoMaterial = If(_cmbMaterial.SelectedItem IsNot Nothing, _cmbMaterial.SelectedItem.ToString(), "")

                Case "GUIA"
                    s.UltimaGuiaMaterial = If(_cmbMaterial.SelectedItem IsNot Nothing, _cmbMaterial.SelectedItem.ToString(), "")
            End Select
        End If

        AddinSettings.Save()

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

End Class
