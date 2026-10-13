Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms
Imports Inventor

''' <summary>
''' Motor de negocio nativo VB.NET para la Selección Base de Piezas (IPT).
''' Reemplaza la regla iLogic "Seleccion baseV2.iLogicVb" e incorpora gestión avanzada de:
''' - Configuración de tipos (CHAPA, PLANCHA, CUADRADO, REDONDO, GUIA, GENERICO)
''' - Generación y cálculo de propiedades MATERIAL_BASE y MATERIAL_PLANO
''' - Inyección dinámica de nuevas reglas de chapa (espesores en mm o pulgadas)
''' - Inyección dinámica de nuevos diámetros de redondo y medidas de cuadrado
''' - Inyección y persistencia de nuevos materiales en las librerías "METALES DB" y "PLASTICOS DB"
''' - Lógica anti-regeneración para preservar desarrollos planos existentes (FlatPattern)
''' </summary>
Public Class SeleccionBaseEngine

    Private ReadOnly _invApp As Inventor.Application
    Private ReadOnly _log As List(Of String)

    ' Subtipos nativos de Autodesk Inventor
    Public Const SUBTYPE_SHEET_METAL As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
    Public Const SUBTYPE_STANDARD_PART As String = "{4D29B490-49B2-11D0-93C3-7E0706000000}"

    Public Sub New(invApp As Inventor.Application)
        _invApp = invApp
        _log = New List(Of String)
    End Sub

    ' =========================================================
    ' EJECUCIÓN PRINCIPAL
    ' =========================================================

    Public Function EjecutarDesdeDocumentoActivo() As Boolean
        Dim targetDoc As Document = Nothing

        ' 1. Edición in situ activa dentro de un ensamblaje (ActiveEditDocument es la pieza .ipt)
        If _invApp.ActiveEditDocument IsNot Nothing AndAlso _invApp.ActiveEditDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            targetDoc = _invApp.ActiveEditDocument

        ' 2. Entorno directo de pieza (.ipt como ActiveDocument)
        ElseIf _invApp.ActiveDocument IsNot Nothing AndAlso _invApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            targetDoc = _invApp.ActiveDocument

        ' 3. Entorno de ensamblaje (.iam): comprobar si el usuario tiene una ocurrencia de pieza seleccionada en el árbol o visor
        ElseIf _invApp.ActiveDocument IsNot Nothing AndAlso _invApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            If _invApp.ActiveSelectSet IsNot Nothing AndAlso _invApp.ActiveSelectSet.Count = 1 Then
                Try
                    Dim selObj As Object = _invApp.ActiveSelectSet.Item(1)
                    If TypeOf selObj Is ComponentOccurrence Then
                        Dim occ As ComponentOccurrence = DirectCast(selObj, ComponentOccurrence)
                        If occ.Definition.Document.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                            targetDoc = DirectCast(occ.Definition.Document, Document)
                        End If
                    End If
                Catch
                End Try
            End If
        End If

        If targetDoc Is Nothing Then
            MessageBox.Show("Selección Base requiere una pieza (.ipt) activa, en edición in situ o seleccionada dentro del ensamblaje.", "Selección Base", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        Return Ejecutar(targetDoc)
    End Function

    Public Function Ejecutar(doc As Document) As Boolean
        Try
            If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
                MessageBox.Show("Selección Base solo puede ejecutarse en documentos de pieza (.ipt).", "Selección Base", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            ' 1. Desplegar el diálogo interactivo moderno para configurar parámetros
            Dim tipo As String = ""
            Dim subTipo As String = ""
            Dim dimEstilo As String = ""
            Dim mat As String = ""
            Dim esMm As Boolean = True

            Using dlg As New SeleccionBaseDialog(_invApp, Me, doc)
                Dim res As DialogResult = dlg.ShowDialog()
                If res <> DialogResult.OK Then Return False

                tipo = dlg.TipoPiezaSeleccionado
                subTipo = dlg.SubTipoFabricacionSeleccionado
                dimEstilo = dlg.DimensionOEstiloSeleccionado
                mat = dlg.MaterialSeleccionado
                esMm = dlg.UnidadesMilimetros
            End Using

            ' 2. El diálogo ya cerró por completo: Inventor recupera el 100% del foco nativo en pantalla.
            ' Si es REDONDO o CUADRADO, solicitar el plano con el cartelito sobre el mouse EXACTAMENTE como Seleccion baseV2.iLogicVb
            Dim oPlanoDestino As Object = Nothing
            Dim generarGeo As Boolean = False

            If tipo = "REDONDO" OrElse tipo = "CUADRADO" Then
                generarGeo = True
                Try
                    ' Pick interactivo nativo de Inventor: genera el cartelito flotante al lado del mouse
                    oPlanoDestino = _invApp.CommandManager.Pick(
                        SelectionFilterEnum.kWorkPlaneFilter,
                        "Seleccione plano de aplicación")
                Catch
                End Try

                ' Si cancela con ESC o no selecciona, asignamos el Plano XY Base (Item 3) como respaldo seguro
                If oPlanoDestino Is Nothing Then
                    Try
                        Dim partDef As PartComponentDefinition = CType(doc, PartDocument).ComponentDefinition
                        oPlanoDestino = partDef.WorkPlanes.Item(3)
                    Catch
                    End Try
                End If
            End If

            ' 3. Aplicar configuración en la pieza (iProperties, Material, y Generar Geometría en el plano seleccionado)
            Dim exito As Boolean = AplicarConfiguracion(
                doc, tipo, subTipo, dimEstilo, mat, esMm,
                generarGeo, False, oPlanoDestino)

            Return exito

        Catch ex As Exception
            MessageBox.Show("Error durante la ejecución de Selección Base:" & vbCrLf & ex.Message, "Error - Selección Base", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' =========================================================
    ' =========================================================
    ' APLICAR CONFIGURACIÓN SELECCIONADA EN LA PIEZA
    ' =========================================================
    Public Function AplicarConfiguracion(
        doc As Document,
        tipoPieza As String,
        subTipoFabricacion As String,
        dimensionOEstilo As String,
        materialSeleccionado As String,
        esMilimetros As Boolean,
        Optional generarGeometria As Boolean = True,
        Optional preferirBocetoNativo As Boolean = True,
        Optional oPlanoDestino As Object = Nothing) As Boolean

        Try
            Dim customProps As PropertySet = GetOrCreatePropertySet(doc, "Inventor User Defined Properties")
            Dim partDoc As PartDocument = DirectCast(doc, PartDocument)
            Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
            Dim oMaterialAsset As AssetLibrary = Nothing
            Dim addData As String = tipoPieza

            Select Case tipoPieza.ToUpper()
                ' =========================================================
                ' 1) CHAPA
                ' =========================================================
                Case "CHAPA"
                    ActivarLibreria("METALES DB", oMaterialAsset)

                    ' Cambiar subtipo a Sheet Metal si no lo es ya
                    If doc.SubType <> SUBTYPE_SHEET_METAL Then
                        doc.SubType = SUBTYPE_SHEET_METAL
                    End If

                    Dim smDef As SheetMetalComponentDefinition = DirectCast(doc.ComponentDefinition, SheetMetalComponentDefinition)
                    Dim debeCambiarEstilo As Boolean = True

                    ' ==== LÓGICA ANTI-REGENERACIÓN ====
                    If smDef.HasFlatPattern AndAlso smDef.FlatPattern IsNot Nothing Then
                        Try
                            If smDef.ActiveSheetMetalStyle IsNot Nothing AndAlso
                               smDef.ActiveSheetMetalStyle.Name.Equals(dimensionOEstilo, StringComparison.OrdinalIgnoreCase) Then
                                 debeCambiarEstilo = False ' Ya coincide, no tocar para no regenerar el desarrollo
                            End If
                        Catch
                        End Try
                    End If

                    If debeCambiarEstilo Then
                        ActivarEstiloChapa(smDef, dimensionOEstilo)
                    End If

                    ' Asignar Material
                    If Not String.IsNullOrEmpty(materialSeleccionado) Then
                        AsignarMaterial(partDoc, materialSeleccionado, oMaterialAsset)
                    End If

                    ' Resolver propiedades
                    addData = "CHAPA"
                    Dim propMatPlano As Inventor.Property = GetOrCreateProperty(customProps, "MATERIAL_PLANO")
                    propMatPlano.Value = "=<MATERIAL_BASE> <SHEET METAL RULE> <MATERIAL>"

                    ' Asegurar que no quede la marca PLASTIC
                    EliminarPropiedadSiExiste(customProps, "PLASTIC")

                ' =========================================================
                ' 2) PLANCHA
                ' =========================================================
                Case "PLANCHA"
                    ActivarLibreria("PLASTICOS DB", oMaterialAsset)

                    If doc.SubType <> SUBTYPE_SHEET_METAL Then
                        doc.SubType = SUBTYPE_SHEET_METAL
                    End If

                    Dim smDef As SheetMetalComponentDefinition = DirectCast(doc.ComponentDefinition, SheetMetalComponentDefinition)
                    ActivarEstiloChapa(smDef, dimensionOEstilo)

                    If Not String.IsNullOrEmpty(materialSeleccionado) Then
                        AsignarMaterial(partDoc, materialSeleccionado, oMaterialAsset)
                    End If

                    addData = "PLANCHA"
                    Dim propMatPlano As Inventor.Property = GetOrCreateProperty(customProps, "MATERIAL_PLANO")
                    propMatPlano.Value = "=<MATERIAL_BASE> <SHEET METAL RULE> <MATERIAL>"

                    ' Marcar PLASTIC
                    Dim propPlastic As Inventor.Property = GetOrCreateProperty(customProps, "PLASTIC")
                    propPlastic.Value = "✓"

                ' =========================================================
                ' 3) CUADRADO
                ' =========================================================
                Case "CUADRADO"
                    If subTipoFabricacion.Equals("PLASTICO", StringComparison.OrdinalIgnoreCase) Then
                        ActivarLibreria("PLASTICOS DB", oMaterialAsset)
                        Dim propPlastic As Inventor.Property = GetOrCreateProperty(customProps, "PLASTIC")
                        propPlastic.Value = "✓"
                    Else
                        ActivarLibreria("METALES DB", oMaterialAsset)
                        EliminarPropiedadSiExiste(customProps, "PLASTIC")
                    End If

                    If doc.SubType <> SUBTYPE_STANDARD_PART Then
                        doc.SubType = SUBTYPE_STANDARD_PART
                    End If

                    If Not String.IsNullOrEmpty(materialSeleccionado) Then
                        AsignarMaterial(partDoc, materialSeleccionado, oMaterialAsset)
                    End If

                    addData = "CUADRADO " & subTipoFabricacion & " " & dimensionOEstilo & " "
                    Dim propMatPlano As Inventor.Property = GetOrCreateProperty(customProps, "MATERIAL_PLANO")
                    propMatPlano.Value = "=<MATERIAL_BASE> <MATERIAL>"

                    ' Inyectar geometría (Boceto nativo o iFeature)
                    If generarGeometria AndAlso Not String.IsNullOrWhiteSpace(dimensionOEstilo) Then
                        GenerarGeometriaMacizo(partDoc, "CUADRADO", dimensionOEstilo, preferirBocetoNativo, oPlanoDestino)
                    End If

                ' =========================================================
                ' 4) REDONDO
                ' =========================================================
                Case "REDONDO"
                    If subTipoFabricacion.Equals("PLASTICO", StringComparison.OrdinalIgnoreCase) Then
                        ActivarLibreria("PLASTICOS DB", oMaterialAsset)
                        Dim propPlastic As Inventor.Property = GetOrCreateProperty(customProps, "PLASTIC")
                        propPlastic.Value = "✓"
                    Else
                        EliminarPropiedadSiExiste(customProps, "PLASTIC")
                        ActivarLibreria("METALES DB", oMaterialAsset)
                    End If

                    If doc.SubType <> SUBTYPE_STANDARD_PART Then
                        doc.SubType = SUBTYPE_STANDARD_PART
                    End If

                    If Not String.IsNullOrEmpty(materialSeleccionado) Then
                        AsignarMaterial(partDoc, materialSeleccionado, oMaterialAsset)
                    End If

                    addData = "REDONDO " & subTipoFabricacion & " " & dimensionOEstilo & " "
                    Dim propMatPlano As Inventor.Property = GetOrCreateProperty(customProps, "MATERIAL_PLANO")
                    propMatPlano.Value = "=<MATERIAL_BASE> <MATERIAL>"

                    ' Inyectar geometría (Boceto nativo o iFeature)
                    If generarGeometria AndAlso Not String.IsNullOrWhiteSpace(dimensionOEstilo) Then
                        GenerarGeometriaMacizo(partDoc, "REDONDO", dimensionOEstilo, preferirBocetoNativo, oPlanoDestino)
                    End If

                ' =========================================================
                ' 5) GENÉRICO
                ' =========================================================
                Case "GENERICO"
                    ActivarLibreria("DEFAULT", oMaterialAsset)
                    If doc.SubType <> SUBTYPE_STANDARD_PART Then
                        doc.SubType = SUBTYPE_STANDARD_PART
                    End If

                    partDoc.ComponentDefinition.Material = ObtenerMaterialPorDefecto(partDoc)
                    addData = "GENERICO"

                    Dim propMatPlano As Inventor.Property = GetOrCreateProperty(customProps, "MATERIAL_PLANO")
                    propMatPlano.Value = "=<MATERIAL_BASE> <MATERIAL>"
                    EliminarPropiedadSiExiste(customProps, "PLASTIC")

                ' =========================================================
                ' 6) GUÍA (Referencia en BOM)
                ' =========================================================
                Case "GUIA"
                    ActivarLibreria("DEFAULT", oMaterialAsset)
                    If doc.SubType <> SUBTYPE_STANDARD_PART Then
                        doc.SubType = SUBTYPE_STANDARD_PART
                    End If

                    partDoc.ComponentDefinition.Material = ObtenerMaterialPorDefecto(partDoc)
                    doc.ComponentDefinition.BOMStructure = BOMStructureEnum.kReferenceBOMStructure

                    addData = "GUIA"
                    Dim propMatPlano As Inventor.Property = GetOrCreateProperty(customProps, "MATERIAL_PLANO")
                    propMatPlano.Value = "=<MATERIAL_BASE> <MATERIAL>"
                    EliminarPropiedadSiExiste(customProps, "PLASTIC")

            End Select

            ' Asignar la propiedad fundamental MATERIAL_BASE
            Dim propMatBase As Inventor.Property = GetOrCreateProperty(customProps, "MATERIAL_BASE")
            propMatBase.Value = addData

            doc.Update()
            Return True

        Catch ex As Exception
            MessageBox.Show("Error aplicando configuración de pieza: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' =========================================================
    ' INYECCIÓN DE NUEVA REGLA DE CHAPA (ESPESOR)
    ' =========================================================
    Public Function InyectarNuevoEspesorChapa(
        doc As Document,
        nombreRegla As String,
        espesorMm As Double) As Boolean

        Try
            If doc.SubType <> SUBTYPE_SHEET_METAL Then
                doc.SubType = SUBTYPE_SHEET_METAL
            End If

            Dim smDef As SheetMetalComponentDefinition = DirectCast(doc.ComponentDefinition, SheetMetalComponentDefinition)

            ' Verificar si ya existe en el documento
            For Each style As SheetMetalStyle In smDef.SheetMetalStyles
                If style.Name.Equals(nombreRegla, StringComparison.OrdinalIgnoreCase) Then
                    style.Activate()
                    Return True
                End If
            Next

            ' Crear una nueva regla a partir de la activa o predeterminada
            Dim baseStyle As SheetMetalStyle = smDef.ActiveSheetMetalStyle
            If baseStyle Is Nothing AndAlso smDef.SheetMetalStyles.Count > 0 Then
                baseStyle = smDef.SheetMetalStyles.Item(1)
            End If

            If baseStyle IsNot Nothing Then
                Dim newStyle As SheetMetalStyle = baseStyle.Copy(nombreRegla)
                ' En Inventor el espesor puede definirse mediante string con unidades
                newStyle.Thickness = espesorMm.ToString(System.Globalization.CultureInfo.InvariantCulture) & " mm"
                newStyle.Activate()
            End If

            ' Persistir en configuración para proyectos futuros y anteriores
            If Not AddinSettings.Current.EspesoresChapaInyectados.Contains(nombreRegla) Then
                AddinSettings.Current.EspesoresChapaInyectados.Add(nombreRegla)
                AddinSettings.Save()
            End If

            Return True
        Catch ex As Exception
            MessageBox.Show("No se pudo inyectar la nueva regla de chapa: " & ex.Message, "Inyección de Regla", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End Try
    End Function

    ' =========================================================
    ' INYECCIÓN DE NUEVO MATERIAL A LA LIBRERÍA Y AL PROYECTO
    ' =========================================================
    Public Function InyectarNuevoMaterial(
        nombreMaterial As String,
        esPlastico As Boolean,
        Optional doc As Document = Nothing) As Boolean

        Try
            Dim nombreLibreria As String = If(esPlastico, "PLASTICOS DB", "METALES DB")
            Dim libAsset As AssetLibrary = ObtenerLibreria(nombreLibreria)

            If libAsset Is Nothing Then
                MessageBox.Show("No se encontró la librería '" & nombreLibreria & "' para guardar el nuevo material." & vbCrLf &
                                "Verifique que el archivo .adsklib se encuentre en 'D:\Bibliotecas\Materiales\' o esté configurado en el proyecto activo.",
                                "Librería no encontrada", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            If libAsset.IsReadOnly Then
                MessageBox.Show("La librería '" & nombreLibreria & "' está protegida contra escritura (Read-Only) en disco." & vbCrLf &
                                "Verifique los atributos del archivo .adsklib en Windows.",
                                "Librería protegida", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            ' Obtener documento de pieza activo o proporcionado para instanciar el Asset
            Dim partDoc As PartDocument = Nothing
            If doc IsNot Nothing AndAlso doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                partDoc = DirectCast(doc, PartDocument)
            ElseIf _invApp.ActiveDocument IsNot Nothing AndAlso _invApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                partDoc = DirectCast(_invApp.ActiveDocument, PartDocument)
            End If

            If partDoc Is Nothing Then
                MessageBox.Show("Se requiere un documento de pieza (.ipt) abierto para generar y registrar el nuevo material en la librería.",
                                "Documento requerido", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            ' 1. Buscar asset base en la propia librería o en el documento para clonar propiedades
            Dim baseAsset As Asset = Nothing
            If libAsset.MaterialAssets.Count > 0 Then
                baseAsset = libAsset.MaterialAssets.Item(1)
            ElseIf partDoc.Assets.Count > 0 Then
                For Each a As Asset In partDoc.Assets
                    If a.AssetType = AssetTypeEnum.kAssetTypeMaterial Then
                        baseAsset = a
                        Exit For
                    End If
                Next
            End If

            If baseAsset Is Nothing Then
                MessageBox.Show("No se encontró ningún material de referencia en la librería para clonar.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            ' 2. Copiar al documento, asignar DisplayName y luego copiar directamente a la biblioteca física
            Dim docAsset As Asset = baseAsset.CopyTo(partDoc)
            docAsset.DisplayName = nombreMaterial.Trim()

            Try
                Dim newLibAsset As Asset = docAsset.CopyTo(libAsset)
            Catch exCopy As Exception
                System.Diagnostics.Debug.WriteLine("Aviso al copiar asset a librería: " & exCopy.Message)
            End Try

            ' 3. Persistir físicamente el archivo .adsklib en disco
            Try
                libAsset.Save()
            Catch exSave As Exception
                MessageBox.Show("No se pudo guardar físicamente el archivo de librería .adsklib: " & exSave.Message, "Aviso de guardado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try

            Return True
        Catch ex As Exception
            MessageBox.Show("Error registrando nuevo material en la librería: " & ex.Message, "Materiales", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End Try
    End Function

    ' =========================================================
    ' OBTENER LIBRERÍA DE FORMA ROBUSTA DESDE INVENTOR O DISCO
    ' =========================================================
    Public Function ObtenerLibreria(nombreLibreria As String, Optional rutaArchivoManual As String = "") As AssetLibrary
        Dim nombreLimpio As String = nombreLibreria.Trim()

        ' 1. Si se especificó una ruta manual de archivo .adsklib
        If Not String.IsNullOrEmpty(rutaArchivoManual) AndAlso IO.File.Exists(rutaArchivoManual) Then
            Try
                Dim libManual As AssetLibrary = _invApp.AssetLibraries.Open(rutaArchivoManual)
                If libManual IsNot Nothing Then
                    _invApp.ActiveMaterialLibrary = libManual
                    Return libManual
                End If
            Catch
            End Try
        End If

        ' 2. Buscar por Item directo en la colección AssetLibraries de la sesión
        Try
            Dim directLib As AssetLibrary = _invApp.AssetLibraries.Item(nombreLimpio)
            If directLib IsNot Nothing Then
                _invApp.ActiveMaterialLibrary = directLib
                Return directLib
            End If
        Catch
        End Try

        ' 3. Buscar por DisplayName o Name ignorando mayúsculas/minúsculas
        For Each libreria As AssetLibrary In _invApp.AssetLibraries
            If libreria.DisplayName.Trim().Equals(nombreLimpio, StringComparison.OrdinalIgnoreCase) OrElse
               libreria.Name.Trim().Equals(nombreLimpio, StringComparison.OrdinalIgnoreCase) Then
                _invApp.ActiveMaterialLibrary = libreria
                Return libreria
            End If
        Next

        ' 4. Buscar en las librerías configuradas en el proyecto activo (ActiveDesignProject.MaterialLibraries)
        Try
            Dim prj = _invApp.DesignProjectManager.ActiveDesignProject
            If prj IsNot Nothing Then
                Dim prjLibs As ProjectAssetLibraries = prj.MaterialLibraries
                For i As Integer = 1 To prjLibs.Count
                    Dim pl As Object = prjLibs.Item(i)
                    Dim libName As String = ""
                    Dim libFile As String = ""
                    Try : libName = pl.Name : Catch : End Try
                    Try : libFile = pl.LibraryFile : Catch : End Try

                    If libName.Equals(nombreLimpio, StringComparison.OrdinalIgnoreCase) OrElse
                       IO.Path.GetFileNameWithoutExtension(libFile).Equals(nombreLimpio, StringComparison.OrdinalIgnoreCase) Then
                        If IO.File.Exists(libFile) Then
                            Try
                                Dim openedPrjLib As AssetLibrary = _invApp.AssetLibraries.Open(libFile)
                                If openedPrjLib IsNot Nothing Then
                                    _invApp.ActiveMaterialLibrary = openedPrjLib
                                    Return openedPrjLib
                                End If
                            Catch
                            End Try
                        End If
                    End If
                Next
            End If
        Catch
        End Try

        ' 5. Buscar en rutas físicas conocidas del sistema
        Dim posiblesRutas As New List(Of String) From {
            IO.Path.Combine("D:\Bibliotecas\Materiales", nombreLimpio & ".adsklib"),
            IO.Path.Combine("D:\Bibliotecas\Materiales", nombreLimpio),
            IO.Path.Combine(AddinSettings.Current.RutaLogs.Replace("RuleManager.log", ""), "Materiales", nombreLimpio & ".adsklib"),
            IO.Path.Combine("C:\ProgramData\Autodesk\Inventor\Material Libraries", nombreLimpio & ".adsklib")
        }

        For Each r In posiblesRutas
            If IO.File.Exists(r) Then
                Try
                    Dim openedLib As AssetLibrary = _invApp.AssetLibraries.Open(r)
                    If openedLib IsNot Nothing Then
                        _invApp.ActiveMaterialLibrary = openedLib
                        Return openedLib
                    End If
                Catch
                End Try
            End If
        Next

        Return Nothing
    End Function

    ' =========================================================
    ' INYECCIÓN DE NUEVO DIÁMETRO (REDONDO)
    ' =========================================================
    Public Sub InyectarNuevoDiametro(nuevoDiametro As String)
        If Not AddinSettings.Current.DiametrosRedondoInyectados.Contains(nuevoDiametro) Then
            AddinSettings.Current.DiametrosRedondoInyectados.Add(nuevoDiametro)
            AddinSettings.Save()
        End If
    End Sub

    ' =========================================================
    ' INYECCIÓN DE NUEVA DIMENSIÓN (CUADRADO)
    ' =========================================================
    Public Sub InyectarNuevaDimensionCuadrado(nuevaDim As String)
        If Not AddinSettings.Current.DimensionesCuadradoInyectados.Contains(nuevaDim) Then
            AddinSettings.Current.DimensionesCuadradoInyectados.Add(nuevaDim)
            AddinSettings.Save()
        End If
    End Sub

    ' =========================================================
    ' LISTAS MAESTRAS DE OPCIONES DE DIMENSIÓN Y REGLAS DE CHAPA
    ' =========================================================

    Public Function ObtenerEspesoresChapaMm() As List(Of String)
        Dim lista As New List(Of String) From {
            "0.3mm", "0.4mm", "0.5mm", "0.8mm", "1mm", "1.2mm", "1.5mm",
            "2mm", "2mm (Nº14)", "2.5mm", "2.95mm", "3mm", "3.4mm", "4mm",
            "5mm", "6mm", "8mm", "10mm", "12mm", "16mm", "20mm", "25mm"
        }
        For Each esp In AddinSettings.Current.EspesoresChapaInyectados
            If Not lista.Contains(esp) Then lista.Add(esp)
        Next
        Return lista
    End Function

    Public Function ObtenerEspesoresChapaPulg() As List(Of String)
        Return New List(Of String) From {
            "1/4"" (6.35mm)", "5/16"" (7.94mm)", "3/8"" (9.53mm)", "1/2"" (12.70mm)",
            "5/8"" (15.88mm)", "3/4"" (19.10mm)", "7/8"" (22.22mm)", "1"" (25.4mm)",
            "1 1/4"" (31.75mm)", "1 1/2"" (38.1mm)"
        }
    End Function

    Public Function ObtenerEspesoresPlancha() As List(Of String)
        Dim lista As New List(Of String) From {
            "2mm", "3mm", "4mm", "6mm", "8mm", "10mm", "12mm", "15mm", "16mm",
            "20mm", "25mm", "30mm", "35mm", "40mm", "50mm", "60mm", "64mm", "65mm"
        }
        For Each esp In AddinSettings.Current.EspesoresPlanchaInyectados
            If Not lista.Contains(esp) Then lista.Add(esp)
        Next
        Return lista
    End Function

    Public Function ObtenerDimensionesCuadrado() As List(Of String)
        Dim lista As New List(Of String) From {
            "3/16"" (4.76mm)", "1/4"" (6.35mm)", "5/16"" (7.94mm)", "3/8"" (9.53mm)", "12mm",
            "1/2"" (12.70mm)", "5/8"" (15.88mm)", "3/4"" (19.10mm)", "20mm", "7/8"" (22.22mm)",
            "1"" (25.4mm)", "30mm", "1 1/4"" (31.75mm)", "1 1/2"" (38.1mm)", "40mm",
            "1 3/4"" (44.45mm)", "45mm", "2"" (50.80mm)", "60mm", "2 1/2"" (63.50mm)",
            "3"" (76.20mm)", "3 1/2"" (88.9mm)"
        }
        For Each d In AddinSettings.Current.DimensionesCuadradoInyectados
            If Not lista.Contains(d) Then lista.Add(d)
        Next
        Return lista
    End Function

    Public Function ObtenerDiametrosRedondo() As List(Of String)
        Dim lista As New List(Of String) From {
            "3/16"" (4.76mm)", "1/4"" (6.35mm)", "5/16"" (7.94mm)", "8mm", "3/8"" (9.53mm)",
            "10mm", "12mm", "1/2"" (12.70mm)", "15mm", "16mm", "5/8"" (15.88mm)", "19mm",
            "3/4"" (19.10mm)", "20mm", "7/8"" (22.22mm)", "25mm", "30mm", "35mm",
            "1 1/4"" (31.75mm)", "1 1/2"" (38.1mm)", "40mm", "1 3/4"" (44.45mm)", "50mm",
            "2"" (50.80mm)", "55mm", "60mm", "2 1/2"" (63.50mm)", "70 mm", "3"" (76.20mm)",
            "3 1/2"" (88.9mm)", "90mm", "110mm", "120mm", "140mm", "150mm", "160mm"
        }
        For Each d In AddinSettings.Current.DiametrosRedondoInyectados
            If Not lista.Contains(d) Then lista.Add(d)
        Next
        Return lista
    End Function

    ' =========================================================
    ' LECTURA ESTRICTA DE MATERIALES DIRECTAMENTE DE LA LIBRERÍA
    ' Sin listas inventadas ni estáticas: lee exactamente lo que hay en la librería física.
    ' Si el usuario añade un material manualmente en Inventor, se leerá directamente de aquí.
    ' =========================================================
    Public Function ObtenerMaterialesDeLibreria(nombreLibreria As String, Optional doc As Document = Nothing, Optional rutaManual As String = "") As List(Of String)
        Dim lista As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        ' 1. Obtener la AssetLibrary real de Autodesk Inventor
        Dim libAsset As AssetLibrary = ObtenerLibreria(nombreLibreria, rutaManual)

        If libAsset IsNot Nothing Then
            Try
                ' Recorrer MaterialAssets de la librería física
                Dim materialsEnum As AssetsEnumerator = libAsset.MaterialAssets
                If materialsEnum IsNot Nothing Then
                    For Each a As Asset In materialsEnum
                        Dim nombre As String = a.DisplayName
                        If String.IsNullOrWhiteSpace(nombre) Then
                            nombre = a.Name
                        End If
                        If Not String.IsNullOrWhiteSpace(nombre) Then
                            lista.Add(nombre.Trim())
                        End If
                    Next
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Error recorriendo MaterialAssets de la librería '" & nombreLibreria & "': " & ex.Message)
            End Try
        End If

        ' 2. Si no se encontraron en la AssetLibrary pero hay un documento activo que ya tiene el material asignado
        If lista.Count = 0 AndAlso doc IsNot Nothing AndAlso doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            Try
                Dim partDoc As PartDocument = DirectCast(doc, PartDocument)
                For Each m As Material In partDoc.Materials
                    If Not String.IsNullOrWhiteSpace(m.Name) Then
                        lista.Add(m.Name.Trim())
                    End If
                Next
            Catch
            End Try
        End If

        ' NOTA IMPORTANTE:
        ' NO se inyectan listas artificiales ni hardcodeadas.
        ' Si la librería está vacía o no se encuentra en disco, se devuelve la lista real (vacía o con lo existente)
        ' para que el usuario pueda ver el estado real y cargar el archivo .adsklib correspondiente o registrar nuevos materiales.

        Dim resultado As List(Of String) = lista.ToList()
        resultado.Sort(StringComparer.OrdinalIgnoreCase)
        Return resultado
    End Function

    ' =========================================================
    ' UTILIDADES PRIVADAS
    ' =========================================================

    Private Sub ActivarLibreria(nombre As String, ByRef oMaterialAsset As AssetLibrary)
        Try
            For Each libreria As AssetLibrary In _invApp.AssetLibraries
                If libreria.DisplayName.Equals(nombre, StringComparison.OrdinalIgnoreCase) Then
                    oMaterialAsset = libreria
                    _invApp.ActiveMaterialLibrary = libreria
                    Exit Sub
                End If
            Next
        Catch
        End Try
    End Sub

    Private Sub ActivarEstiloChapa(smDef As SheetMetalComponentDefinition, nombreEstilo As String)
        Try
            For Each style As SheetMetalStyle In smDef.SheetMetalStyles
                If style.Name.Equals(nombreEstilo, StringComparison.OrdinalIgnoreCase) Then
                    style.Activate()
                    Exit Sub
                End If
            Next
            ' Si no existe, activar el primero disponible
            If smDef.SheetMetalStyles.Count > 0 Then
                smDef.SheetMetalStyles.Item(1).Activate()
            End If
        Catch
        End Try
    End Sub

    Private Sub AsignarMaterial(partDoc As PartDocument, nombreMaterial As String, libAsset As AssetLibrary)
        Try
            ' Buscar en materiales locales del documento
            For Each m As Material In partDoc.Materials
                If m.Name.Equals(nombreMaterial, StringComparison.OrdinalIgnoreCase) Then
                    partDoc.ComponentDefinition.Material = m
                    Exit Sub
                End If
            Next

            ' Intentar copiar desde la biblioteca activa
            If libAsset IsNot Nothing Then
                For Each a As Asset In libAsset.MaterialAssets
                    If a.DisplayName.Equals(nombreMaterial, StringComparison.OrdinalIgnoreCase) Then
                        Dim localAsset As Asset = a.CopyTo(partDoc)
                        Exit Sub
                    End If
                Next
            End If

            ' Fallback mediante iProperties
            Dim designProps As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")
            designProps.Item("Material").Value = nombreMaterial
        Catch
        End Try
    End Sub

    Private Function ObtenerMaterialPorDefecto(partDoc As PartDocument) As Material
        Try
            Return partDoc.Materials.Item("Genérico")
        Catch
            Return partDoc.ComponentDefinition.Material
        End Try
    End Function

    Private Function GetOrCreatePropertySet(doc As Document, setName As String) As PropertySet
        Try
            Return doc.PropertySets.Item(setName)
        Catch
            Return doc.PropertySets.Add(setName)
        End Try
    End Function

    Private Function GetOrCreateProperty(props As PropertySet, propName As String) As Inventor.Property
        For Each p As Inventor.Property In props
            If p.Name.Equals(propName, StringComparison.OrdinalIgnoreCase) Then Return p
        Next
        Return props.Add(" ", propName)
    End Function

    Private Sub EliminarPropiedadSiExiste(props As PropertySet, propName As String)
        Try
            For Each p As Inventor.Property In props
                If p.Name.Equals(propName, StringComparison.OrdinalIgnoreCase) Then
                    p.Delete()
                    Exit Sub
                End If
            Next
        Catch
        End Try
    End Sub

    ' =========================================================
    ' PARSEO Y CONVERSIÓN DE MEDIDAS (REDONDO / CUADRADO)
    ' =========================================================
    Public Shared Function ExtraerMedidaLimpia(texto As String) As String
        If String.IsNullOrWhiteSpace(texto) Then Return "10mm"
        Dim inicio As Integer = texto.IndexOf("("c)
        If inicio >= 0 Then
            Dim fin As Integer = texto.IndexOf(")"c)
            If fin > inicio Then
                Return texto.Substring(inicio + 1, fin - inicio - 1).Trim()
            End If
        End If
        Return texto.Trim()
    End Function

    Public Shared Function ParsearMedidaEnCm(texto As String) As Double
        Dim limp As String = ExtraerMedidaLimpia(texto)
        Dim numSolo As String = limp.ToLower().Replace("mm", "").Trim()
        Dim dVal As Double = 10.0
        If Double.TryParse(numSolo, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, dVal) Then
            Return dVal / 10.0 ' 1 mm = 0.1 cm
        ElseIf Double.TryParse(numSolo, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, dVal) Then
            Return dVal / 10.0
        End If
        Return 1.0 ' 10 mm por defecto
    End Function

    ' =========================================================
    ' GENERACIÓN DE BOCETO PARAMÉTRICO NATIVO: REDONDO
    ' =========================================================
    Public Function GenerarBocetoRedondo(
        partDoc As PartDocument,
        dimensionTexto As String,
        Optional oPlano As Object = Nothing) As Boolean

        Try
            Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
            Dim medidaLimpia As String = ExtraerMedidaLimpia(dimensionTexto)
            Dim expr As String = medidaLimpia
            If Not expr.ToLower().EndsWith("mm") AndAlso Not expr.Contains("""") Then
                expr = expr & " mm"
            End If

            ' 1. Si ya existe el parámetro "DIAMETRO", actualizarlo directamente
            Dim paramExistente As Inventor.Parameter = Nothing
            Try
                paramExistente = compDef.Parameters.Item("DIAMETRO")
            Catch
            End Try

            If paramExistente IsNot Nothing Then
                paramExistente.Expression = expr
                partDoc.Update()
                Return True
            End If

            ' 2. Seleccionar plano de trabajo
            Dim oPlane As Object = oPlano
            If oPlane Is Nothing Then
                oPlane = compDef.WorkPlanes.Item(3) ' Plano XY base
            End If

            ' 3. Crear boceto plano
            Dim oSketch As PlanarSketch = compDef.Sketches.Add(oPlane, False)
            oSketch.Name = "Boceto REDONDO " & medidaLimpia

            Dim tg As TransientGeometry = _invApp.TransientGeometry
            Dim radioCm As Double = ParsearMedidaEnCm(dimensionTexto) / 2.0
            If radioCm <= 0 Then radioCm = 0.5

            ' 4. Crear círculo centrado
            Dim centerPoint As Point2d = tg.CreatePoint2d(0, 0)
            Dim oCircle As SketchCircle = oSketch.SketchCircles.AddByCenterRadius(centerPoint, radioCm)

            ' 5. Restricción coincidente con el origen del boceto
            Try
                oSketch.GeometricConstraints.AddCoincident(oCircle.CenterSketchPoint, oSketch.OriginPoint)
            Catch
            End Try

            ' 6. Cota de diámetro paramétrica
            Try
                Dim textPoint As Point2d = tg.CreatePoint2d(radioCm * 1.3, radioCm * 1.3)
                Dim oDim As DimensionConstraint = oSketch.DimensionConstraints.AddDiameter(oCircle, textPoint)
                oDim.Parameter.Expression = expr
                Try
                    oDim.Parameter.Name = "DIAMETRO"
                Catch
                End Try
            Catch
            End Try

            oSketch.Visible = True
            partDoc.Update()
            Return True

        Catch ex As Exception
            MessageBox.Show("Aviso al generar boceto REDONDO: " & ex.Message, "Geometría", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End Try
    End Function

    ' =========================================================
    ' GENERACIÓN DE BOCETO PARAMÉTRICO NATIVO: CUADRADO
    ' =========================================================
    Public Function GenerarBocetoCuadrado(
        partDoc As PartDocument,
        dimensionTexto As String,
        Optional oPlano As Object = Nothing) As Boolean

        Try
            Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
            Dim medidaLimpia As String = ExtraerMedidaLimpia(dimensionTexto)
            Dim expr As String = medidaLimpia
            If Not expr.ToLower().EndsWith("mm") AndAlso Not expr.Contains("""") Then
                expr = expr & " mm"
            End If

            ' 1. Si ya existe el parámetro "LADO", actualizarlo directamente
            Dim paramExistente As Inventor.Parameter = Nothing
            Try
                paramExistente = compDef.Parameters.Item("LADO")
            Catch
            End Try

            If paramExistente IsNot Nothing Then
                paramExistente.Expression = expr
                partDoc.Update()
                Return True
            End If

            ' 2. Seleccionar plano de trabajo
            Dim oPlane As Object = oPlano
            If oPlane Is Nothing Then
                oPlane = compDef.WorkPlanes.Item(3) ' Plano XY base
            End If

            ' 3. Crear boceto plano
            Dim oSketch As PlanarSketch = compDef.Sketches.Add(oPlane, False)
            oSketch.Name = "Boceto CUADRADO " & medidaLimpia

            Dim tg As TransientGeometry = _invApp.TransientGeometry
            Dim ladoCm As Double = ParsearMedidaEnCm(dimensionTexto)
            If ladoCm <= 0 Then ladoCm = 1.0
            Dim halfCm As Double = ladoCm / 2.0

            ' 4. Trazar 4 líneas cerradas del cuadrado
            Dim pt1 As Point2d = tg.CreatePoint2d(-halfCm, -halfCm)
            Dim pt2 As Point2d = tg.CreatePoint2d(halfCm, -halfCm)
            Dim pt3 As Point2d = tg.CreatePoint2d(halfCm, halfCm)
            Dim pt4 As Point2d = tg.CreatePoint2d(-halfCm, halfCm)

            Dim line1 As SketchLine = oSketch.SketchLines.AddByTwoPoints(pt1, pt2)
            Dim line2 As SketchLine = oSketch.SketchLines.AddByTwoPoints(line1.EndSketchPoint, pt3)
            Dim line3 As SketchLine = oSketch.SketchLines.AddByTwoPoints(line2.EndSketchPoint, pt4)
            Dim line4 As SketchLine = oSketch.SketchLines.AddByTwoPoints(line3.EndSketchPoint, line1.StartSketchPoint)

            ' 5. Restricciones geométricas (horizontal, vertical, perpendicular, igualdad)
            Try
                oSketch.GeometricConstraints.AddHorizontal(line1)
                oSketch.GeometricConstraints.AddVertical(line2)
                oSketch.GeometricConstraints.AddHorizontal(line3)
                oSketch.GeometricConstraints.AddVertical(line4)
                oSketch.GeometricConstraints.AddPerpendicular(line1, line2)
                oSketch.GeometricConstraints.AddEqual(line1, line2)
            Catch
            End Try

            ' 6. Centrado en el origen
            Try
                Dim diag As SketchLine = oSketch.SketchLines.AddByTwoPoints(line1.StartSketchPoint, line2.EndSketchPoint)
                diag.Construction = True
                oSketch.GeometricConstraints.AddMidpoint(oSketch.OriginPoint, diag)
            Catch
            End Try

            ' 7. Cota de dimensión paramétrica
            Try
                Dim textPtX As Point2d = tg.CreatePoint2d(0, -halfCm - 0.4)
                Dim oDimX As DimensionConstraint = oSketch.DimensionConstraints.AddTwoPointDistance(
                    line1.StartSketchPoint, line1.EndSketchPoint, DimensionOrientationEnum.kHorizontalDim, textPtX)
                oDimX.Parameter.Expression = expr
                Try
                    oDimX.Parameter.Name = "LADO"
                Catch
                End Try
            Catch
            End Try

            oSketch.Visible = True
            partDoc.Update()
            Return True

        Catch ex As Exception
            MessageBox.Show("Aviso al generar boceto CUADRADO: " & ex.Message, "Geometría", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End Try
    End Function

    ' =========================================================
    ' INSERCIÓN DE IFEATURE CUSTOM (.ide)
    ' =========================================================
    Public Function InsertariFeatureMacizo(
        partDoc As PartDocument,
        tipo As String,
        dimensionTexto As String,
        Optional oPlano As Object = Nothing) As Boolean

        Try
            Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
            Dim iFeaturePath As String = System.IO.Path.Combine("D:\Bibliotecas\iFeatures\custom", "MACIZO " & tipo & ".ide")

            If Not System.IO.File.Exists(iFeaturePath) Then
                Return False
            End If

            Dim oPlane As Object = oPlano
            If oPlane Is Nothing Then
                oPlane = compDef.WorkPlanes.Item(3) ' Plano XY base
            End If

            Dim numeroStr As String = ExtraerMedidaLimpia(dimensionTexto)

            Dim oFeatures As PartFeatures = compDef.Features
            Dim oiFeatureDef As iFeatureDefinition = oFeatures.iFeatures.CreateiFeatureDefinition(iFeaturePath)

            For Each oInput As iFeatureInput In oiFeatureDef.iFeatureInputs
                Select Case oInput.Name
                    Case "Plano de base"
                        Dim oPlaneInput As iFeatureSketchPlaneInput = TryCast(oInput, iFeatureSketchPlaneInput)
                        If oPlaneInput IsNot Nothing Then
                            oPlaneInput.PlaneInput = oPlane
                        End If
                    Case "Dimension"
                        Dim oParamInput As iFeatureParameterInput = TryCast(oInput, iFeatureParameterInput)
                        If oParamInput IsNot Nothing Then
                            oParamInput.Expression = numeroStr
                        End If
                End Select
            Next

            oFeatures.iFeatures.Add(oiFeatureDef)
            partDoc.Update()
            Return True

        Catch ex As Exception
            ' Si falla la inserción de la iFeature, retornamos False para fallback a Boceto
            Return False
        End Try
    End Function

    ' =========================================================
    ' ENLACE GENERAL DE GEOMETRÍA: BOCETO O IFEATURE
    ' =========================================================
    Public Function GenerarGeometriaMacizo(
        partDoc As PartDocument,
        tipo As String,
        dimensionTexto As String,
        Optional preferirBocetoNativo As Boolean = True,
        Optional oPlano As Object = Nothing) As Boolean

        ' Si NO se prefiere boceto nativo y se solicitó intentar iFeature:
        If Not preferirBocetoNativo Then
            Dim okIFeature As Boolean = InsertariFeatureMacizo(partDoc, tipo, dimensionTexto, oPlano)
            If okIFeature Then Return True
        End If

        ' Crear boceto nativo 100% estable
        If tipo.Equals("REDONDO", StringComparison.OrdinalIgnoreCase) Then
            Return GenerarBocetoRedondo(partDoc, dimensionTexto, oPlano)
        ElseIf tipo.Equals("CUADRADO", StringComparison.OrdinalIgnoreCase) Then
            Return GenerarBocetoCuadrado(partDoc, dimensionTexto, oPlano)
        End If

        Return False
    End Function

End Class
