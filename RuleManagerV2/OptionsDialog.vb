Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Windows.Forms
Imports Inventor
Imports TextBox = System.Windows.Forms.TextBox

''' <summary>
''' Diálogo de Opciones del Add-In RuleManager V2.
''' Estructurado con pestañas temáticas similar a "Opciones de la aplicación" de Autodesk Inventor.
''' </summary>
Public Class OptionsDialog
    Inherits Form

    Private _tabControl As TabControl

    ' Controles - Tab General
    Private _chkCrearCarpetas As CheckBox
    Private _chkSincLibrerias As CheckBox
    Private _chkSilentOp As CheckBox
    Private _txtRutaLogs As TextBox

    ' Controles - Tab Ensamblaje
    Private _chkOmitirCC As CheckBox
    Private _chkDetectarSim As CheckBox
    Private _chkActualizarCant As CheckBox

    ' Controles - Tab Planos (Dibujo)
    Private _chkAutoListaPiezas As CheckBox
    Private _cmbEstiloPieza As ComboBox
    Private _cmbEstiloEnsamblaje As ComboBox
    Private _cmbUbicacionListaPiezas As ComboBox
    Private _cmbNivelBom As ComboBox
    Private _cmbFormatoPdf As ComboBox
    Private _chkOrdenarHojas As CheckBox
    Private _txtPlantillaExcel As TextBox

    ' Controles - Tab Pieza
    Private _chkCuboCorte As CheckBox
    Private _chkChapa As CheckBox

    ' Controles - Tab Proyecto & Archivos
    Private _chkExcluirOldVersions As CheckBox
    Private _chkPreservarEspeciales As CheckBox
    Private _chkProtegerRouted As CheckBox
    Private _chkEnrutarMulticorte As CheckBox

    ' Controles - Tab Odoo ERP
    Private _cmbTipoBom As ComboBox
    Private _txtPrefijoExtId As TextBox

    ' Controles - Tab Backup & Red
    Private _txtRutaLocalBib As TextBox
    Private _txtRutaServidorBib As TextBox
    Private _cmbIntervaloBackup As ComboBox
    Private _chkAutoSyncBib As CheckBox

    ' Botones inferiores
    Private _btnAceptar As Button
    Private _btnCancelar As Button
    Private _btnAplicar As Button
    Private _btnImportar As Button
    Private _btnExportar As Button

    Private ReadOnly _invApp As Inventor.Application

    Public Sub New(Optional invApp As Inventor.Application = Nothing)
        If invApp IsNot Nothing Then
            _invApp = invApp
        Else
            Try
                _invApp = DirectCast(Microsoft.VisualBasic.Interaction.GetObject(Nothing, "Inventor.Application"), Inventor.Application)
            Catch
                _invApp = Nothing
            End Try
        End If

        ' --- Ventana ---
        Text = "RuleManager - Opciones de la aplicación"
        StartPosition = FormStartPosition.CenterScreen
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False
        MinimizeBox = False
        ShowInTaskbar = False
        ClientSize = New Size(620, 500)
        Font = New Font("Segoe UI", 9.0F)

        InitializeComponent()
        CargarValores()
    End Sub

    Private Sub InitializeComponent()
        ' --- TabControl Principal ---
        _tabControl = New TabControl With {
            .Location = New System.Drawing.Point(10, 10),
            .Size = New Size(595, 435),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        }

        ' Crear Pestañas
        Dim tabGeneral As New TabPage("General")
        Dim tabEnsamblaje As New TabPage("Ensamblaje")
        Dim tabPlanos As New TabPage("Dibujo")
        Dim tabPieza As New TabPage("Pieza")
        Dim tabProyecto As New TabPage("Proyecto & Archivos")
        Dim tabOdoo As New TabPage("Odoo ERP")
        Dim tabBackup As New TabPage("Backup & Red")

        ConstruirTabGeneral(tabGeneral)
        ConstruirTabEnsamblaje(tabEnsamblaje)
        ConstruirTabPlanos(tabPlanos)
        ConstruirTabPieza(tabPieza)
        ConstruirTabProyecto(tabProyecto)
        ConstruirTabOdoo(tabOdoo)
        ConstruirTabBackup(tabBackup)

        _tabControl.TabPages.Add(tabGeneral)
        _tabControl.TabPages.Add(tabEnsamblaje)
        _tabControl.TabPages.Add(tabPlanos)
        _tabControl.TabPages.Add(tabPieza)
        _tabControl.TabPages.Add(tabProyecto)
        _tabControl.TabPages.Add(tabOdoo)
        _tabControl.TabPages.Add(tabBackup)

        Controls.Add(_tabControl)

        ' --- Botones Inferiores ---
        Dim yBtn As Integer = 455

        _btnImportar = New Button With {
            .Text = "Importar...",
            .Location = New System.Drawing.Point(12, yBtn),
            .Size = New Size(85, 28)
        }

        _btnExportar = New Button With {
            .Text = "Exportar...",
            .Location = New System.Drawing.Point(102, yBtn),
            .Size = New Size(85, 28)
        }

        _btnAceptar = New Button With {
            .Text = "Aceptar",
            .Location = New System.Drawing.Point(350, yBtn),
            .Size = New Size(80, 28)
        }
        AddHandler _btnAceptar.Click, AddressOf OnAceptar

        _btnCancelar = New Button With {
            .Text = "Cancelar",
            .DialogResult = DialogResult.Cancel,
            .Location = New System.Drawing.Point(435, yBtn),
            .Size = New Size(80, 28)
        }

        _btnAplicar = New Button With {
            .Text = "Aplicar",
            .Location = New System.Drawing.Point(520, yBtn),
            .Size = New Size(80, 28)
        }
        AddHandler _btnAplicar.Click, AddressOf OnAplicar

        Controls.Add(_btnImportar)
        Controls.Add(_btnExportar)
        Controls.Add(_btnAceptar)
        Controls.Add(_btnCancelar)
        Controls.Add(_btnAplicar)

        AcceptButton = _btnAceptar
        CancelButton = _btnCancelar
    End Sub

    Private Sub ConstruirTabGeneral(tab As TabPage)
        tab.Controls.Add(CrearEncabezadoSeccion("Comportamiento de Inicio y Proyectos", 15, 12))

        _chkCrearCarpetas = New CheckBox With {
            .Text = "Crear estructura de carpetas (01 - DIBUJOS, 02 - PLANOS, etc.) al cambiar de proyecto",
            .Location = New System.Drawing.Point(20, 35),
            .Size = New Size(510, 25)
        }

        _chkSincLibrerias = New CheckBox With {
            .Text = "Sincronizar bibliotecas y Content Center automáticamente (StartupTasks)",
            .Location = New System.Drawing.Point(20, 65),
            .Size = New Size(510, 25)
        }

        _chkSilentOp = New CheckBox With {
            .Text = "Habilitar modo silencioso (SilentOperation) para evitar alertas y diálogos bloqueantes",
            .Location = New System.Drawing.Point(20, 95),
            .Size = New Size(510, 25)
        }

        tab.Controls.Add(CrearEncabezadoSeccion("Rutas de Auditoría y Diagnóstico", 15, 135))

        Dim lblLogs As New Label With {.Text = "Ruta de archivo log:", .Location = New System.Drawing.Point(20, 160), .AutoSize = True}
        _txtRutaLogs = New TextBox With {.Location = New System.Drawing.Point(160, 157), .Size = New Size(360, 23)}

        tab.Controls.Add(_chkCrearCarpetas)
        tab.Controls.Add(_chkSincLibrerias)
        tab.Controls.Add(_chkSilentOp)
        tab.Controls.Add(lblLogs)
        tab.Controls.Add(_txtRutaLogs)
    End Sub

    Private Sub ConstruirTabEnsamblaje(tab As TabPage)
        tab.Controls.Add(CrearEncabezadoSeccion("Componentes Comerciales (COM)", 15, 12))

        _chkOmitirCC = New CheckBox With {
            .Text = "Omitir componentes de Content Center y bulonería normalizada al asignar COM",
            .Location = New System.Drawing.Point(20, 35),
            .Size = New Size(510, 25)
        }

        tab.Controls.Add(CrearEncabezadoSeccion("Conteo y Simetrías", 15, 75))

        _chkDetectarSim = New CheckBox With {
            .Text = "Detección recursiva automática de componentes espejados (SEARCHSIMETRIAS)",
            .Location = New System.Drawing.Point(20, 100),
            .Size = New Size(510, 25)
        }

        _chkActualizarCant = New CheckBox With {
            .Text = "Actualizar propiedad CANTIDAD_USADA en todos los componentes referenciados",
            .Location = New System.Drawing.Point(20, 130),
            .Size = New Size(510, 25)
        }

        tab.Controls.Add(_chkOmitirCC)
        tab.Controls.Add(_chkDetectarSim)
        tab.Controls.Add(_chkActualizarCant)
    End Sub

    Private Sub ConstruirTabPlanos(tab As TabPage)
        ' SECCIÓN 1: Inserción Automática de Lista de Piezas
        tab.Controls.Add(CrearEncabezadoSeccion("Inserción Automática de Lista de Piezas (Parts List)", 15, 10))

        _chkAutoListaPiezas = New CheckBox With {
            .Text = "Insertar lista de piezas automáticamente al colocar la primera vista en una hoja",
            .Location = New System.Drawing.Point(20, 32),
            .Size = New Size(540, 22),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(15, 23, 42)
        }

        Dim lblEstiloPieza As New Label With {
            .Text = "Lista de piezas para Piezas (.ipt):",
            .Location = New System.Drawing.Point(20, 62),
            .Size = New Size(220, 20),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        _cmbEstiloPieza = New ComboBox With {
            .Location = New System.Drawing.Point(245, 59),
            .Size = New Size(285, 23),
            .DropDownStyle = ComboBoxStyle.DropDown
        }

        Dim lblEstiloEnsamblaje As New Label With {
            .Text = "Lista de piezas para Ensamblajes (.iam):",
            .Location = New System.Drawing.Point(20, 92),
            .Size = New Size(220, 20),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        _cmbEstiloEnsamblaje = New ComboBox With {
            .Location = New System.Drawing.Point(245, 89),
            .Size = New Size(285, 23),
            .DropDownStyle = ComboBoxStyle.DropDown
        }

        Dim lblUbicacion As New Label With {
            .Text = "Ubicación en la hoja de dibujo:",
            .Location = New System.Drawing.Point(20, 122),
            .Size = New Size(220, 20),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        _cmbUbicacionListaPiezas = New ComboBox With {
            .Location = New System.Drawing.Point(245, 119),
            .Size = New Size(285, 23),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        _cmbUbicacionListaPiezas.Items.AddRange(New Object() {
            "Esquina superior derecha (Border MaxPoint)",
            "Sobre el rótulo / cajetín (TitleBlock)",
            "Esquina superior izquierda",
            "Esquina inferior izquierda"
        })

        Dim lblNivelBom As New Label With {
            .Text = "Nivel BoM para Ensamblajes:",
            .Location = New System.Drawing.Point(20, 152),
            .Size = New Size(220, 20),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        _cmbNivelBom = New ComboBox With {
            .Location = New System.Drawing.Point(245, 149),
            .Size = New Size(285, 23),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        _cmbNivelBom.Items.AddRange(New Object() {
            "Estructurado (Primer nivel)",
            "Sólo piezas (Parts only)"
        })

        ' Poblar estilos estándar conocidos y estilos dinámicos del documento activo
        CargarEstilosDisponibles()

        ' Habilitar / deshabilitar controles según el CheckBox
        AddHandler _chkAutoListaPiezas.CheckedChanged, Sub()
                                                           Dim en As Boolean = _chkAutoListaPiezas.Checked
                                                           lblEstiloPieza.Enabled = en
                                                           _cmbEstiloPieza.Enabled = en
                                                           lblEstiloEnsamblaje.Enabled = en
                                                           _cmbEstiloEnsamblaje.Enabled = en
                                                           lblUbicacion.Enabled = en
                                                           _cmbUbicacionListaPiezas.Enabled = en
                                                           lblNivelBom.Enabled = en
                                                           _cmbNivelBom.Enabled = en
                                                       End Sub

        ' SECCIÓN 2: Listas de Materiales y Extracción
        tab.Controls.Add(CrearEncabezadoSeccion("Listas de Materiales y Extracción a Excel", 15, 185))

        Dim lblTmpl As New Label With {.Text = "Plantilla Excel LDM:", .Location = New System.Drawing.Point(20, 212), .AutoSize = True}
        _txtPlantillaExcel = New TextBox With {.Location = New System.Drawing.Point(160, 209), .Size = New Size(370, 23)}

        _chkOrdenarHojas = New CheckBox With {
            .Text = "Ordenar hojas de dibujo alfabética y numéricamente en documentos .idw",
            .Location = New System.Drawing.Point(20, 240),
            .Size = New Size(510, 25)
        }

        ' SECCIÓN 3: Impresión a PDF
        tab.Controls.Add(CrearEncabezadoSeccion("Impresión a PDF", 15, 275))

        Dim lblPdf As New Label With {.Text = "Formato PDF predeterminado:", .Location = New System.Drawing.Point(20, 302), .AutoSize = True}
        _cmbFormatoPdf = New ComboBox With {
            .Location = New System.Drawing.Point(200, 299),
            .Size = New Size(180, 23),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        _cmbFormatoPdf.Items.AddRange(New Object() {"B3", "B3_EXT", "LEGAL"})

        tab.Controls.AddRange(New Control() {
            _chkAutoListaPiezas,
            lblEstiloPieza, _cmbEstiloPieza,
            lblEstiloEnsamblaje, _cmbEstiloEnsamblaje,
            lblUbicacion, _cmbUbicacionListaPiezas,
            lblNivelBom, _cmbNivelBom,
            lblTmpl, _txtPlantillaExcel,
            _chkOrdenarHojas,
            lblPdf, _cmbFormatoPdf
        })
    End Sub

    Private Sub CargarEstilosDisponibles()
        ' Estilos estándar basados en plantillas de la empresa
        Dim estilosSugeridosPieza As New List(Of String) From {
            "Copia de PIEZAS",
            "PIEZAS",
            "LDM",
            "Lista de piezas (ISO)"
        }

        Dim estilosSugeridosEnsamblaje As New List(Of String) From {
            "Copia de ENSAMBLAJE",
            "Copia de Copia de ENSAMBLAJE",
            "LDMARM",
            "LDM",
            "BULONERIA",
            "INSUMOS",
            "Lista de materiales (ISO)",
            "Lista de piezas (ISO)"
        }

        ' Si hay un documento de dibujo activo o abierto en Inventor, leer los estilos reales de StylesManager
        Try
            Dim drawDoc As DrawingDocument = Nothing
            If _invApp IsNot Nothing Then
                If _invApp.ActiveDocument IsNot Nothing AndAlso
                   _invApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                    drawDoc = DirectCast(_invApp.ActiveDocument, DrawingDocument)
                End If
            End If

            If drawDoc IsNot Nothing AndAlso drawDoc.StylesManager IsNot Nothing Then
                For Each s As PartsListStyle In drawDoc.StylesManager.PartsListStyles
                    Dim sName As String = s.Name.Trim()
                    If Not estilosSugeridosPieza.Contains(sName, StringComparer.OrdinalIgnoreCase) Then
                        estilosSugeridosPieza.Add(sName)
                    End If
                    If Not estilosSugeridosEnsamblaje.Contains(sName, StringComparer.OrdinalIgnoreCase) Then
                        estilosSugeridosEnsamblaje.Add(sName)
                    End If
                Next
            End If
        Catch
        End Try

        _cmbEstiloPieza.Items.Clear()
        For Each item In estilosSugeridosPieza
            _cmbEstiloPieza.Items.Add(item)
        Next

        _cmbEstiloEnsamblaje.Items.Clear()
        For Each item In estilosSugeridosEnsamblaje
            _cmbEstiloEnsamblaje.Items.Add(item)
        Next
    End Sub

    Private Sub ConstruirTabPieza(tab As TabPage)
        tab.Controls.Add(CrearEncabezadoSeccion("Cálculos y Dimensionado de Piezas", 15, 12))

        _chkCuboCorte = New CheckBox With {
            .Text = "Calcular automáticamente cubo de corte (LARGO, ANCHO, ESPESOR) en DimensionPart",
            .Location = New System.Drawing.Point(20, 35),
            .Size = New Size(510, 25)
        }

        _chkChapa = New CheckBox With {
            .Text = "Detectar piezas de chapa metálica y calcular desarrollo / factor K (PB CHAPA)",
            .Location = New System.Drawing.Point(20, 65),
            .Size = New Size(510, 25)
        }

        tab.Controls.Add(_chkCuboCorte)
        tab.Controls.Add(_chkChapa)
    End Sub

    Private Sub ConstruirTabProyecto(tab As TabPage)
        tab.Controls.Add(CrearEncabezadoSeccion("Reglas de Clasificación y Renombrado", 15, 12))

        _chkExcluirOldVersions = New CheckBox With {
            .Text = "Excluir carpetas 'OldVersions' en cualquier nivel de la ruta",
            .Location = New System.Drawing.Point(20, 35),
            .Size = New Size(510, 25)
        }

        _chkPreservarEspeciales = New CheckBox With {
            .Text = "Preservar subcarpetas especiales de ensamblaje (AIP y Frame) dentro de 01 - DIBUJOS",
            .Location = New System.Drawing.Point(20, 65),
            .Size = New Size(510, 25)
        }

        _chkProtegerRouted = New CheckBox With {
            .Text = "Proteger sistemas ruteados (conduits, tuberías, cables) durante el renombrado",
            .Location = New System.Drawing.Point(20, 95),
            .Size = New Size(510, 25)
        }

        _chkEnrutarMulticorte = New CheckBox With {
            .Text = "Enrutar archivos DXF/DWG con nombre 'MULTICORTE' hacia '03 - CORTES'",
            .Location = New System.Drawing.Point(20, 125),
            .Size = New Size(510, 25)
        }

        tab.Controls.Add(_chkExcluirOldVersions)
        tab.Controls.Add(_chkPreservarEspeciales)
        tab.Controls.Add(_chkProtegerRouted)
        tab.Controls.Add(_chkEnrutarMulticorte)
    End Sub

    Private Sub ConstruirTabOdoo(tab As TabPage)
        tab.Controls.Add(CrearEncabezadoSeccion("Parámetros de Exportación Odoo ERP", 15, 12))

        Dim lblTipoBom As New Label With {.Text = "Tipo de BoM predeterminado:", .Location = New System.Drawing.Point(20, 38), .AutoSize = True}
        _cmbTipoBom = New ComboBox With {
            .Location = New System.Drawing.Point(200, 35),
            .Size = New Size(160, 23),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        _cmbTipoBom.Items.AddRange(New Object() {"Kit", "Manufacture"})

        Dim lblPrefijo As New Label With {.Text = "Prefijo External ID (Proyecto):", .Location = New System.Drawing.Point(20, 70), .AutoSize = True}
        _txtPrefijoExtId = New TextBox With {.Location = New System.Drawing.Point(200, 67), .Size = New Size(160, 23)}

        tab.Controls.Add(lblTipoBom)
        tab.Controls.Add(_cmbTipoBom)
        tab.Controls.Add(lblPrefijo)
        tab.Controls.Add(_txtPrefijoExtId)
    End Sub

    Private Sub ConstruirTabBackup(tab As TabPage)
        Dim secBackup As Control = CrearEncabezadoSeccion("Sincronización y Backup de Bibliotecas (Local ⇄ Servidor)", 15, 15)

        Dim lblLocal As New Label With {.Text = "Ruta Local (PC):", .Location = New System.Drawing.Point(20, 48), .AutoSize = True}
        _txtRutaLocalBib = New TextBox With {.Location = New System.Drawing.Point(170, 45), .Size = New Size(320, 23)}

        Dim lblServidor As New Label With {.Text = "Ruta Servidor Red:", .Location = New System.Drawing.Point(20, 80), .AutoSize = True}
        _txtRutaServidorBib = New TextBox With {.Location = New System.Drawing.Point(170, 77), .Size = New Size(320, 23)}

        Dim lblIntervalo As New Label With {.Text = "Intervalo Programado:", .Location = New System.Drawing.Point(20, 114), .AutoSize = True}
        _cmbIntervaloBackup = New ComboBox With {
            .Location = New System.Drawing.Point(170, 111),
            .Size = New Size(160, 23),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        _cmbIntervaloBackup.Items.AddRange(New Object() {"Manual", "Diario", "Semanal", "Mensual"})

        _chkAutoSyncBib = New CheckBox With {
            .Text = "Sincronizar automáticamente al iniciar Inventor si se venció el intervalo",
            .Location = New System.Drawing.Point(20, 150),
            .Size = New Size(500, 22),
            .AutoSize = True
        }

        Dim btnAbrirSyncDialog As New Button With {
            .Text = "Abrir Administrador de Backup / Sincronización...",
            .Location = New System.Drawing.Point(20, 185),
            .Size = New Size(320, 28),
            .BackColor = System.Drawing.Color.FromArgb(241, 245, 249),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnAbrirSyncDialog.Click, Sub()
                                                 Using dlg As New BackupSyncDialog(_invApp)
                                                     dlg.ShowDialog(Me)
                                                 End Using
                                             End Sub

        tab.Controls.AddRange({secBackup, lblLocal, _txtRutaLocalBib, lblServidor, _txtRutaServidorBib, lblIntervalo, _cmbIntervaloBackup, _chkAutoSyncBib, btnAbrirSyncDialog})
    End Sub

    Private Function CrearEncabezadoSeccion(titulo As String, x As Integer, y As Integer) As Control
        Dim pnl As New Panel With {.Location = New System.Drawing.Point(x, y), .Size = New Size(520, 20)}
        Dim lbl As New Label With {
            .Text = titulo,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = System.Drawing.Color.FromArgb(0, 50, 120),
            .AutoSize = True,
            .Location = New System.Drawing.Point(0, 0)
        }
        pnl.Controls.Add(lbl)
        Return pnl
    End Function

    Private Sub CargarValores()
        Dim s = AddinSettings.Current
        _chkCrearCarpetas.Checked = s.CrearCarpetasProyecto
        _chkSincLibrerias.Checked = s.SincronizarLibreriasAuto
        _chkSilentOp.Checked = s.HabilitarSilentOperation
        _txtRutaLogs.Text = s.RutaLogs

        _chkOmitirCC.Checked = s.OmitirContentCenterCOM
        _chkDetectarSim.Checked = s.DetectarSimetriasAuto
        _chkActualizarCant.Checked = s.ActualizarCantidadUsada

        ' Planos / Lista de piezas
        _chkAutoListaPiezas.Checked = s.AutoInsertarListaPiezas
        _cmbEstiloPieza.Text = If(Not String.IsNullOrEmpty(s.EstiloListaPiezasPieza), s.EstiloListaPiezasPieza, "Copia de PIEZAS")
        _cmbEstiloEnsamblaje.Text = If(Not String.IsNullOrEmpty(s.EstiloListaPiezasEnsamblaje), s.EstiloListaPiezasEnsamblaje, "Copia de ENSAMBLAJE")

        Select Case s.UbicacionListaPiezas
            Case "SobreRotulo"
                _cmbUbicacionListaPiezas.SelectedIndex = 1
            Case "SuperiorIzquierda"
                _cmbUbicacionListaPiezas.SelectedIndex = 2
            Case "InferiorIzquierda"
                _cmbUbicacionListaPiezas.SelectedIndex = 3
            Case Else
                _cmbUbicacionListaPiezas.SelectedIndex = 0
        End Select

        Select Case s.NivelBomEnsamblaje
            Case "SoloPiezas"
                _cmbNivelBom.SelectedIndex = 1
            Case Else
                _cmbNivelBom.SelectedIndex = 0
        End Select

        _cmbFormatoPdf.SelectedItem = s.FormatoPdfPredeterminado
        _chkOrdenarHojas.Checked = s.OrdenarHojasAuto
        _txtPlantillaExcel.Text = s.PlantillaLdmExcel

        _chkCuboCorte.Checked = s.CalcularCuboCorteAuto
        _chkChapa.Checked = s.ExtraerParametrosChapa

        _chkExcluirOldVersions.Checked = s.ExcluirOldVersions
        _chkPreservarEspeciales.Checked = s.PreservarCarpetasEspeciales
        _chkProtegerRouted.Checked = s.ProtegerRoutedSystems
        _chkEnrutarMulticorte.Checked = s.EnrutarMulticorteDxf

        _cmbTipoBom.SelectedItem = s.TipoBomPredeterminado
        _txtPrefijoExtId.Text = s.PrefijoExternalId

        _txtRutaLocalBib.Text = If(Not String.IsNullOrEmpty(s.RutaLocalBibliotecas), s.RutaLocalBibliotecas, LibraryBackupSyncManager.DEFAULT_LOCAL_PATH)
        _txtRutaServidorBib.Text = If(Not String.IsNullOrEmpty(s.RutaServidorBibliotecas), s.RutaServidorBibliotecas, LibraryBackupSyncManager.DEFAULT_SERVER_PATH)
        _cmbIntervaloBackup.SelectedItem = If(Not String.IsNullOrEmpty(s.IntervaloBackupBibliotecas), s.IntervaloBackupBibliotecas, "Diario")
        _chkAutoSyncBib.Checked = s.AutoSyncBibliotecasAlIniciar
    End Sub

    Private Sub GuardarValores()
        Dim s = AddinSettings.Current
        s.CrearCarpetasProyecto = _chkCrearCarpetas.Checked
        s.SincronizarLibreriasAuto = _chkSincLibrerias.Checked
        s.HabilitarSilentOperation = _chkSilentOp.Checked
        s.RutaLogs = _txtRutaLogs.Text.Trim()

        s.OmitirContentCenterCOM = _chkOmitirCC.Checked
        s.DetectarSimetriasAuto = _chkDetectarSim.Checked
        s.ActualizarCantidadUsada = _chkActualizarCant.Checked

        ' Planos / Lista de piezas
        s.AutoInsertarListaPiezas = _chkAutoListaPiezas.Checked
        s.EstiloListaPiezasPieza = _cmbEstiloPieza.Text.Trim()
        s.EstiloListaPiezasEnsamblaje = _cmbEstiloEnsamblaje.Text.Trim()

        Select Case _cmbUbicacionListaPiezas.SelectedIndex
            Case 1
                s.UbicacionListaPiezas = "SobreRotulo"
            Case 2
                s.UbicacionListaPiezas = "SuperiorIzquierda"
            Case 3
                s.UbicacionListaPiezas = "InferiorIzquierda"
            Case Else
                s.UbicacionListaPiezas = "SuperiorDerecha"
        End Select

        Select Case _cmbNivelBom.SelectedIndex
            Case 1
                s.NivelBomEnsamblaje = "SoloPiezas"
            Case Else
                s.NivelBomEnsamblaje = "Estructurado"
        End Select

        If _cmbFormatoPdf.SelectedItem IsNot Nothing Then
            s.FormatoPdfPredeterminado = _cmbFormatoPdf.SelectedItem.ToString()
        End If
        s.OrdenarHojasAuto = _chkOrdenarHojas.Checked
        s.PlantillaLdmExcel = _txtPlantillaExcel.Text.Trim()

        s.CalcularCuboCorteAuto = _chkCuboCorte.Checked
        s.ExtraerParametrosChapa = _chkChapa.Checked

        s.ExcluirOldVersions = _chkExcluirOldVersions.Checked
        s.PreservarCarpetasEspeciales = _chkPreservarEspeciales.Checked
        s.ProtegerRoutedSystems = _chkProtegerRouted.Checked
        s.EnrutarMulticorteDxf = _chkEnrutarMulticorte.Checked

        If _cmbTipoBom.SelectedItem IsNot Nothing Then
            s.TipoBomPredeterminado = _cmbTipoBom.SelectedItem.ToString()
        End If
        s.PrefijoExternalId = _txtPrefijoExtId.Text.Trim()

        s.RutaLocalBibliotecas = _txtRutaLocalBib.Text.Trim()
        s.RutaServidorBibliotecas = _txtRutaServidorBib.Text.Trim()
        If _cmbIntervaloBackup.SelectedItem IsNot Nothing Then
            s.IntervaloBackupBibliotecas = _cmbIntervaloBackup.SelectedItem.ToString()
        End If
        s.AutoSyncBibliotecasAlIniciar = _chkAutoSyncBib.Checked

        AddinSettings.Save()
    End Sub

    Private Sub OnAplicar(sender As Object, e As EventArgs)
        GuardarValores()
    End Sub

    Private Sub OnAceptar(sender As Object, e As EventArgs)
        GuardarValores()
        DialogResult = DialogResult.OK
        Close()
    End Sub

End Class
