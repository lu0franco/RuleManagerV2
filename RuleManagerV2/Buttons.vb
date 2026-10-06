Imports Inventor
Imports System.IO
Imports System.Drawing
Imports System.Windows.Forms

Public Class Buttons

    Private _inventor As Inventor.Application

    ' =========================
    ' DATOS DE BOTONES
    ' =========================

    Private _buttonsData As New List(Of ButtonData)

    ' =========================
    ' BOTONES REALES INVENTOR
    ' =========================

    Private _buttons As New Dictionary(
        Of String, ButtonDefinition)

    ' =========================
    ' ACCIONES DE BOTONES
    ' =========================

    Private _buttonActions As New Dictionary(
        Of String, Action(Of Inventor.Application))


    ' =========================================================
    ' CONSTRUCTOR
    ' =========================================================

    Public Sub New(inventor As Inventor.Application)

        _inventor = inventor

        LoadButtonsData()

        CreateButtons()

    End Sub


    ' =========================================================
    ' CARGAR CONFIGURACION BOTONES
    ' =========================================================
    Private Sub LoadButtonsData()

        ' =====================================================
        ' ASM
        ' =====================================================

        _buttonsData.Add(New ButtonData(
            "Asignar Comerciales",
            "BTN_COM_ASIGN",
            "Utilice esta herramienta para indicar los comerciales en el ensamblaje.",
            "Assembly",
            AddressOf Rules.EjecutarSComAsign,
            "icon_asm_asignar_comerciales_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "Asignar Propiedades",
            "BTN_PROP_ASIGN",
            "Esta herramienta revisará ensamblajes y piezas agregando Propiedades Personalizadas...",
            "Assembly",
            AddressOf Rules.EjecutarPropAsign,
            "icon_asm_asignar_propiedades_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "Buscar Simetrías",
            "BTN_SIM",
            "Esta herramienta detectará simetrías utilizando la referencia de almacén...",
            "Assembly",
            AddressOf Rules.EjecutarSim,
            "icon_asm_sim_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "Contador de partes",
            "BTN_PARTS_COUNT",
            "Utilice esta herramienta para contar las partes del ensamblaje...",
            "Assembly",
            AddressOf Rules.EjecutarPartsCount,
            "icon_asm_parts_count_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "Revisión de simetrías",
            "BTN_REV_SIM",
            "Utilice esta herramienta en caso de experimentar un comportamiento extraño...",
            "Assembly",
            AddressOf Rules.EjecutarRevSim,
            "icon_asm_rev_sim_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "Convertir Proyecto Definitivo",
            "BTN_PROY_RENAME",
            "Convierte un proyecto preliminar en definitivo usando datos del ensamblaje principal.",
            "Assembly",
            AddressOf Rules.EjecutarProjectRename,
            "icon_ordenar_proyecto_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "Renombrar Componentes",
            "BTN_RENOMBRAR_COMP",
            "Renombra automáticamente las piezas y ensamblajes del modelo según Part Number y Stock Number.",
            "Assembly",
            AddressOf Rules.EjecutarRenombrarComponentes,
            "icon_ordenar_proyecto_32x32.png"))

        ' =====================================================
        ' IPT
        ' =====================================================

        _buttonsData.Add(New ButtonData(
            "1) SELECCION BASE",
            "BTN_SEL_BASE",
            "SELECCION BASE",
            "Part",
            AddressOf Rules.EjecutarSelBase,
            "icon_ipt_sel_base_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "2) PROP ASIGN PART",
            "BTN_PROP_PART",
            "PROP ASIGN PART",
            "Part",
            AddressOf Rules.EjecutarPropAsignPart,
            "icon_ipt_prop_asign_32x32.png"))


        ' =====================================================
        ' IDW
        ' =====================================================

        _buttonsData.Add(New ButtonData(
            "Traer Cantidades",
            "BTN_TRAER_CANT",
            "Utilice esta herramienta para obtener las propiedades de CANTIDAD_USADA...",
            "Drawing",
            AddressOf Rules.EjecutarTraerCantidades,
            "icon_1_traer_cantidades_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "Añadir Copias",
            "BTN_ADD_COPIAS",
            "Utilice esta herramienta para añadir las hojas de copia de trabajo...",
            "Drawing",
            AddressOf Rules.EjecutarAddCopias,
            "icon_2_add_copias_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "Eliminar Copias",
            "BTN_DEL_COPIAS",
            "Utilice esta herramienta para eliminar las hojas de copia de trabajo...",
            "Drawing",
            AddressOf Rules.EjecutarDelCopias,
            "icon_3_del_copias_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "Cambiar Referencias",
            "BTN_CAMBIAR_REF",
            "Utilice esta herramienta si desea cambiar las referencias de una hoja de dibujo...",
            "Drawing",
            AddressOf Rules.EjecutarCambiarReferencias,
            "icon_4_cambiar_ref_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "Ordenar Planos",
            "BTN_ORD_PLANOS",
            "Utilice esta herramienta para ordenar los planos alfabéticamente...",
            "Drawing",
            AddressOf Rules.EjecutarOrdPlanos,
            "icon_5_ordenar_planos_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "LDM Manager",
            "BTN_LDM_MANAGER",
            "Utilice esta herramienta para extraer la Lista de Materiales de todos los planos a EXCEL...",
            "Drawing",
            AddressOf Rules.EjecutarLDMMANAGER,
            "icon_ldm_manager_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "Eliminar Listas",
            "BTN_DEL_LISTAS",
            "Utilice esta herramienta en caso de que alguna extracción de LDM MANAGER falle...",
            "Drawing",
            AddressOf Rules.EjecutarDelListas,
            "icon_6_eliminar_listas_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "Legales a PDF",
            "BTN_PRINT_LEGALES",
            "Utilice esta herramienta para extraer en PDF las hojas en FORMATO LEGAL...",
            "Drawing",
            AddressOf Rules.EjecutarPrintLEGALes,
            "icon_7_print_legales_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "B3 a PDF",
            "BTN_PRINT_B3",
            "Utilice esta herramienta para extraer en PDF las hojas EN FORMATO B3...",
            "Drawing",
            AddressOf Rules.EjecutarPrintB3,
            "icon_8_print_b3_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "B3 EXT a PDF",
            "BTN_PRINT_B3_EXT",
            "Utilice esta herramienta para extraer en PDF las hojas EN FORMATO B3 EXTENDIDO...",
            "Drawing",
            AddressOf Rules.EjecutarPrintB3EXT,
            "icon_9_print_b3_ext_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "B3 EXT WIDE a PDF",
            "BTN_PRINT_B3_WIDE",
            "Utilice esta herramienta para extraer en PDF las hojas EN FORMATO B3 EXTENDIDO WIDE...",
            "Drawing",
            AddressOf Rules.EjecutarPrintB3EXTWIDE,
            "icon_10_print_b3_ext_wide_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "Organizar Proyecto",
            "BTN_ORG_PROY",
            "Clasifica automáticamente archivos del proyecto según reglas internas.",
            "ZeroDoc",
            AddressOf Rules.EjecutarOrganizarProyecto,
            "icon_10_print_b3_ext_wide_32x32.png"))

        _buttonsData.Add(New ButtonData(
            "OPCIONES",
            "BTN_OPCIONES",
            "Configura el comportamiento del addin RuleManager.",
            "ZeroDoc",
            AddressOf Rules.EjecutarOpciones,
            "icon_10_print_b3_ext_wide_32x32.png"))

    End Sub


    ' =========================================================
    ' CREAR BOTONES
    ' =========================================================

    Private Sub CreateButtons()

        Dim conDefs As ControlDefinitions =
            _inventor.CommandManager.ControlDefinitions

        For Each btnData In _buttonsData

            Dim btn As ButtonDefinition

            btn = conDefs.AddButtonDefinition(
                btnData.DisplayName,
                btnData.InternalName,
                CommandTypesEnum.kEditMaskCmdType,
                System.Guid.NewGuid().ToString(),
                btnData.Description,
                btnData.Description,
                Nothing,
                LoadIcon(btnData.LargeIcon))

            ' Crear un handler específico que capture el InternalName
            Dim internalName As String = btnData.InternalName
            AddHandler btn.OnExecute, Sub() ExecuteButtonAction(internalName)

            _buttons.Add(
                btnData.InternalName,
                btn)

            _buttonActions.Add(
                btnData.InternalName,
                btnData.ExecuteAction)

            AddButtonToRibbon(
                btn,
                btnData)

        Next

    End Sub


    ' =========================================================
    ' AGREGAR BOTON A RIBBON
    ' =========================================================

    Private Sub AddButtonToRibbon(
        btn As ButtonDefinition,
        btnData As ButtonData)

        Dim ribbon As Ribbon =
            _inventor.UserInterfaceManager.Ribbons.Item(
                btnData.Environment)

        Dim ribbonTab As RibbonTab

        ' =====================================================
        ' TAB
        ' =====================================================

        Try

            ribbonTab =
                ribbon.RibbonTabs.Item("TAB_TECNICA")

        Catch

            ribbonTab =
                ribbon.RibbonTabs.Add(
                    "TECNICA",
                    "TAB_TECNICA",
                    System.Guid.NewGuid().ToString())

        End Try


        ' =====================================================
        ' PANEL
        ' =====================================================

        Dim ribbonPanel As RibbonPanel

        Try

            ribbonPanel =
                ribbonTab.RibbonPanels.Item(
                    "PANEL_AUTOMATIZACION")

        Catch

            ribbonPanel =
                ribbonTab.RibbonPanels.Add(
                    "AUTOMATIZACION",
                    "PANEL_AUTOMATIZACION",
                    System.Guid.NewGuid().ToString())

        End Try


        ' =====================================================
        ' AGREGAR BOTON
        ' =====================================================

        ribbonPanel.CommandControls.AddButton(
            btn,
            True)

    End Sub


    ' =========================================================
    ' CLICK GENERICO
    ' =========================================================

    Private Sub ExecuteButtonAction(internalName As String)

        Try

            If _buttonActions.ContainsKey(internalName) Then

                _buttonActions(internalName).Invoke(_inventor)

            End If

        Catch ex As Exception

            MessageBox.Show(ex.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub




    Private Function LoadIcon(
        fileName As String) As Object

        Dim addinPath As String =
            System.IO.Path.GetDirectoryName(
                Reflection.Assembly.
                GetExecutingAssembly().
                Location)

        Dim iconPath As String =
            System.IO.Path.Combine(
                addinPath,
                "Icons",
                fileName)

        If Not System.IO.File.Exists(iconPath) Then
            MessageBox.Show("No existe: " & iconPath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return Nothing
        End If

        Dim bmp As New Bitmap(iconPath)
        Return IconManager.ToPictureDisp(bmp)

    End Function

End Class
