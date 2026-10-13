Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

''' <summary>
''' Motor de conteo recursivo de componentes y detección de simetrías (IAM).
''' Traduce y optimiza la regla iLogic 'Contador de elementosV2.iLogicVb' para el Add-In RuleManager V2.
''' 
''' 1. Detecta piezas y subensamblajes simétricos (por patrón espejo o sufijo 'MIR').
''' 2. Normaliza la descripción a 'SIMETRIA' y restaura el Part Number base para agrupar izquierda y derecha.
''' 3. Cuenta recursivamente todas las ocurrencias activas por Part Number + Stock Number.
''' 4. Inyecta en las iProperties personalizadas:
'''    - CANTIDAD_USADA: total de componentes en el ensamblaje.
'''    - CANTIDAD_DERECHA: cantidad de piezas normales/derechas.
'''    - CANTIDAD_IZQUIERDA: cantidad de piezas simétricas/espejadas.
'''    - TEXTO_CANTIDADES: '=<CANTIDAD_DERECHA> DER.  <CANTIDAD_IZQUIERDA> IZQ.' o ' ' si no es mixta.
''' </summary>
Public Class ContadorElementosEngine

    Private ReadOnly _invApp As Inventor.Application

    Public Sub New(invApp As Inventor.Application)
        _invApp = invApp
    End Sub

    ''' <summary>
    ''' Ejecuta el flujo completo de conteo y detección de simetrías.
    ''' </summary>
    Public Sub Ejecutar()
        Try
            ' 1. Validar entorno de ensamblaje activo
            If _invApp.ActiveDocument Is Nothing Then
                MessageBox.Show("No hay ningún documento activo en Autodesk Inventor.",
                                "Contador de Elementos", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If _invApp.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Esta herramienta solo funciona en documentos de ensamblaje (.iam).",
                                "Contador de Elementos - Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If

            Dim asmDoc As AssemblyDocument = DirectCast(_invApp.ActiveDocument, AssemblyDocument)

            ' =========================================================
            ' PASO 1: DETECTAR Y MARCAR SIMETRÍAS (RECURSIVO)
            ' =========================================================
            Dim simetricasLog As New List(Of String)
            Dim listaSimetricas As New List(Of ComponentOccurrence)

            FindMirrorOccurrencesRecursive(asmDoc.ComponentDefinition.Occurrences, simetricasLog, listaSimetricas)

            ' Procesar las simetrías encontradas: Asignar Description='SIMETRIA' y corregir Part Number base
            Dim procesadosSim As New Dictionary(Of String, Boolean)
            Dim nombresSimetriasUnicos As New List(Of String)

            For Each occ As ComponentOccurrence In listaSimetricas
                Try
                    Dim doc As Document = occ.Definition.Document
                    If doc Is Nothing Then Continue For

                    Dim keyDoc As String = GetDocKey(doc)
                    Dim nombreMostrar As String = doc.DisplayName

                    If Not nombresSimetriasUnicos.Contains(nombreMostrar) Then
                        nombresSimetriasUnicos.Add(nombreMostrar)
                    End If

                    If procesadosSim.ContainsKey(keyDoc) Then Continue For

                    ' Marcar descripción como SIMETRIA en Design Tracking Properties
                    Try
                        Dim stdProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                        stdProps.Item("Description").Value = "SIMETRIA"
                    Catch
                    End Try

                    ' Obtener Part Number original (quitando el espejo) para que agrupe con la derecha
                    Dim originalPN As String = GetOriginalPartNumberFromOccurrence(occ)
                    If Not String.IsNullOrEmpty(originalPN) Then
                        Try
                            Dim stdProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                            stdProps.Item("Part Number").Value = originalPN
                        Catch
                        End Try
                    End If

                    procesadosSim(keyDoc) = True
                Catch
                End Try
            Next

            ' =========================================================
            ' PASO 2: CONTEO RECURSIVO DE OCURRENCIAS
            ' =========================================================
            Dim totalCounts As New Dictionary(Of String, Integer)
            Dim simCounts As New Dictionary(Of String, Integer)
            Dim docList As New Dictionary(Of String, List(Of Document))

            CountSubAssemblies(asmDoc.ComponentDefinition.Occurrences, totalCounts, simCounts, docList)

            ' =========================================================
            ' PASO 3: ESCRITURA DE CANTIDADES EN TODAS LAS PIEZAS
            ' =========================================================
            Dim totalDocumentosActualizados As Integer = AssignToAllParts(totalCounts, simCounts, docList)

            ' =========================================================
            ' PASO 4: REPORTE FINAL AL USUARIO
            ' =========================================================
            Dim mensajeFinal As String =
                "Proceso de conteo completado con éxito." & vbCrLf & vbCrLf &
                "• Grupos de componentes analizados: " & totalCounts.Count & vbCrLf &
                "• Archivos actualizados con iProperties: " & totalDocumentosActualizados & vbCrLf & vbCrLf

            If nombresSimetriasUnicos.Count > 0 Then
                mensajeFinal &= "Se detectaron y actualizaron las siguientes simetrías:" & vbCrLf &
                                "-------------------------------------------------------" & vbCrLf &
                                String.Join(vbCrLf, nombresSimetriasUnicos)

                MessageBox.Show(mensajeFinal, "Éxito - Simetrías Identificadas y Cantidades Actualizadas",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                mensajeFinal &= "No se detectaron piezas o subensamblajes con simetría en ningún nivel." & vbCrLf &
                                "Todas las cantidades generales (CANTIDAD_USADA) fueron actualizadas correctamente."

                MessageBox.Show(mensajeFinal, "Éxito - Conteo de Partes Finalizado",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("Error durante la ejecución del Contador de Elementos:" & vbCrLf & ex.Message,
                            "Contador de Elementos - Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' =========================================================================
    ' RECURSIVIDAD PASO 1: BUSCAR SIMETRÍAS
    ' =========================================================================
    Private Sub FindMirrorOccurrencesRecursive(
        occurrences As ComponentOccurrences,
        ByRef simetricasLog As List(Of String),
        ByRef listaSimetricas As List(Of ComponentOccurrence))

        If occurrences Is Nothing Then Return

        For Each occ As ComponentOccurrence In occurrences
            Try
                If occ.Suppressed Then Continue For
                If occ.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then Continue For
                If Not occ.Visible Then Continue For

                ' Saltar miembros de contenido estándar (Content Center)
                If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Try
                        Dim pDef As PartComponentDefinition = TryCast(occ.Definition, PartComponentDefinition)
                        If pDef IsNot Nothing AndAlso pDef.IsContentMember Then
                            Continue For
                        End If
                    Catch
                    End Try
                End If

                Dim isMirror As Boolean = False
                Dim pn As String = ""

                ' 1) Evaluar si es simetría por patrón de espejo
                If occ.PatternElement IsNot Nothing Then
                    Try
                        Dim patElemObj As Object = occ.PatternElement
                        Dim compObj As Object = patElemObj.Components
                        If compObj IsNot Nothing Then
                            Dim compTypeName As String = compObj.Type.ToString().ToUpper()
                            ' kMirrorPatternType = 111874 o detección textual
                            If compTypeName.Contains("MIRROR") OrElse compTypeName.Contains("SIMETR") OrElse CInt(compObj.Type) = 111874 Then
                                isMirror = True
                            End If
                        End If
                    Catch
                        Try
                            Dim parentPat As Object = occ.PatternElement.Parent
                            If parentPat IsNot Nothing Then
                                Dim patName As String = parentPat.Name.ToString().ToUpper()
                                If patName.Contains("MIRROR") OrElse patName.Contains("SIMETR") OrElse patName.Contains("ESPEJO") Then
                                    isMirror = True
                                End If
                            End If
                        Catch
                        End Try
                    End Try
                End If

                ' 2) Evaluar por PN que contenga "MIR"
                If Not isMirror Then
                    Try
                        Dim doc As Document = occ.Definition.Document
                        If doc IsNot Nothing Then
                            Dim stdProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                            pn = stdProps.Item("Part Number").Value.ToString()
                            If pn.ToUpper().Contains("MIR") Then
                                isMirror = True
                            End If
                        End If
                    Catch
                    End Try
                End If

                ' Si es simetría, la guardamos
                If isMirror Then
                    simetricasLog.Add(occ.Name)
                    listaSimetricas.Add(occ)
                End If

                ' LLAMADA RECURSIVA: Ir a buscar dentro de los subensamblajes
                If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Try
                        If occ.SubOccurrences IsNot Nothing AndAlso occ.SubOccurrences.Count > 0 Then
                            FindMirrorOccurrencesRecursive(occ.SubOccurrences, simetricasLog, listaSimetricas)
                        End If
                    Catch
                    End Try
                End If
            Catch
            End Try
        Next
    End Sub

    ' =========================================================================
    ' RECURSIVIDAD PASO 2: CONTAR COMPONENTES
    ' =========================================================================
    Private Sub CountSubAssemblies(
        occurrences As ComponentOccurrences,
        ByRef totalCounts As Dictionary(Of String, Integer),
        ByRef simCounts As Dictionary(Of String, Integer),
        ByRef docList As Dictionary(Of String, List(Of Document)))

        If occurrences Is Nothing Then Return

        For Each oOcc As ComponentOccurrence In occurrences
            Try
                If oOcc.Suppressed Then Continue For
                If oOcc.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then Continue For
                If Not oOcc.Visible Then Continue For

                ' Omitir Content Center
                If oOcc.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Try
                        Dim pDef As PartComponentDefinition = TryCast(oOcc.Definition, PartComponentDefinition)
                        If pDef IsNot Nothing AndAlso pDef.IsContentMember Then
                            Continue For
                        End If
                    Catch
                    End Try
                End If

                ' LLAMADA RECURSIVA: Baja a los niveles inferiores
                If oOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Try
                        If oOcc.SubOccurrences IsNot Nothing AndAlso oOcc.SubOccurrences.Count > 0 Then
                            CountSubAssemblies(oOcc.SubOccurrences, totalCounts, simCounts, docList)
                        End If
                    Catch
                    End Try
                End If

                ' Contabilizar ocurrencia actual
                Try
                    Dim doc As Document = oOcc.Definition.Document
                    If doc Is Nothing Then Continue For

                    Dim props As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                    Dim nombre As String = GetPropertyValue(props, "Part Number").Trim()
                    Dim stock As String = GetPropertyValue(props, "Stock Number").Trim()
                    Dim desc As String = GetPropertyValue(props, "Description").Trim()

                    If String.IsNullOrEmpty(nombre) Then nombre = "Sin nombre"
                    If String.IsNullOrEmpty(stock) Then stock = "Sin número de almacenamiento"

                    Dim key As String = nombre & "_" & stock

                    ' Contar total general
                    If totalCounts.ContainsKey(key) Then
                        totalCounts(key) += 1
                    Else
                        totalCounts.Add(key, 1)
                    End If

                    ' Contar simetrías
                    If desc.ToUpper() = "SIMETRIA" Then
                        If simCounts.ContainsKey(key) Then
                            simCounts(key) += 1
                        Else
                            simCounts.Add(key, 1)
                        End If
                    End If

                    ' Guardar documento en lista sin duplicar referencia de archivo
                    If docList.ContainsKey(key) Then
                        Dim yaExiste As Boolean = False
                        Dim docKeyActual As String = GetDocKey(doc)
                        For Each d As Document In docList(key)
                            If GetDocKey(d) = docKeyActual Then
                                yaExiste = True
                                Exit For
                            End If
                        Next

                        If Not yaExiste Then
                            docList(key).Add(doc)
                        End If
                    Else
                        docList.Add(key, New List(Of Document) From {doc})
                    End If

                Catch
                End Try
            Catch
            End Try
        Next
    End Sub

    ' =========================================================================
    ' PROCESADO FINAL: ASIGNAR QUANTITIES A IPROPERTIES
    ' =========================================================================
    Private Function AssignToAllParts(
        totalCounts As Dictionary(Of String, Integer),
        simCounts As Dictionary(Of String, Integer),
        docList As Dictionary(Of String, List(Of Document))) As Integer

        Dim docsActualizados As Integer = 0

        For Each key As String In docList.Keys
            Dim total As Integer = If(totalCounts.ContainsKey(key), totalCounts(key), 0)
            Dim sim As Integer = If(simCounts.ContainsKey(key), simCounts(key), 0)
            Dim derechas As Integer = total - sim

            For Each doc As Document In docList(key)
                Try
                    Dim userProps As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")

                    ' Asignar propiedades personalizadas requeridas
                    SetOrCreateProperty(userProps, "CANTIDAD_USADA", total)
                    SetOrCreateProperty(userProps, "CANTIDAD_DERECHA", derechas)
                    SetOrCreateProperty(userProps, "CANTIDAD_IZQUIERDA", sim)

                    If (sim <> 0 AndAlso derechas <> 0) Then
                        SetOrCreateProperty(userProps, "TEXTO_CANTIDADES", "=<CANTIDAD_DERECHA> DER.  <CANTIDAD_IZQUIERDA> IZQ.")
                    Else
                        SetOrCreateProperty(userProps, "TEXTO_CANTIDADES", " ")
                    End If

                    docsActualizados += 1
                Catch
                End Try
            Next
        Next

        Return docsActualizados
    End Function

    ' =========================================================================
    ' HELPERS (FUNCIONES AUXILIARES)
    ' =========================================================================

    ''' <summary>
    ''' Recupera el Part Number original del componente base descartando el estado espejo.
    ''' </summary>
    Private Function GetOriginalPartNumberFromOccurrence(occ As ComponentOccurrence) As String
        Try
            Dim doc As Document = occ.Definition.Document
            If doc Is Nothing Then Return ""

            Dim pnActual As String = GetSafePN(doc)

            ' --- PIEZA: intentar mediante componente derivado ---
            If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                Try
                    Dim pDoc As PartDocument = DirectCast(doc, PartDocument)
                    Dim refComps As ReferenceComponents = pDoc.ComponentDefinition.ReferenceComponents

                    If refComps IsNot Nothing AndAlso refComps.DerivedPartComponents.Count > 0 Then
                        For i As Integer = 1 To refComps.DerivedPartComponents.Count
                            Try
                                Dim dpc As DerivedPartComponent = refComps.DerivedPartComponents.Item(i)
                                Dim origDoc As Document = dpc.ReferencedDocumentDescriptor.ReferencedDocument
                                Dim pn As String = GetSafePN(origDoc)
                                If Not String.IsNullOrEmpty(pn) Then Return pn
                            Catch
                            End Try
                        Next
                    End If
                Catch
                End Try
            End If

            ' --- ENSAMBLE: intentar derivado de ensamblaje ---
            If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                Try
                    Dim aDoc As AssemblyDocument = DirectCast(doc, AssemblyDocument)
                    Dim refCompsAsm As ReferenceComponents = aDoc.ComponentDefinition.ReferenceComponents

                    If refCompsAsm IsNot Nothing AndAlso refCompsAsm.DerivedAssemblyComponents.Count > 0 Then
                        For i As Integer = 1 To refCompsAsm.DerivedAssemblyComponents.Count
                            Try
                                Dim dac As DerivedAssemblyComponent = refCompsAsm.DerivedAssemblyComponents.Item(i)
                                Dim origDoc As Document = dac.ReferencedDocumentDescriptor.ReferencedDocument
                                Dim pn As String = GetSafePN(origDoc)
                                If Not String.IsNullOrEmpty(pn) Then Return pn
                            Catch
                            End Try
                        Next
                    End If
                Catch
                End Try
            End If

            ' --- Fallback 1: Buscar en documentos referenciados sin 'MIR' ---
            Try
                If doc.ReferencedDocuments.Count > 0 Then
                    Dim candidato As String = ""
                    For Each r As Document In doc.ReferencedDocuments
                        Dim pnRef As String = GetSafePN(r)
                        If Not String.IsNullOrEmpty(pnRef) Then
                            If Not pnRef.ToUpper().Contains("MIR") Then
                                Return pnRef
                            End If
                            If String.IsNullOrEmpty(candidato) Then candidato = pnRef
                        End If
                    Next
                    If Not String.IsNullOrEmpty(candidato) Then Return candidato
                End If
            Catch
            End Try

            ' --- Fallback 2: Limpieza de texto eliminando el 'MIR' ---
            If Not String.IsNullOrEmpty(pnActual) Then
                Return StripMIR(pnActual)
            End If

        Catch
        End Try

        Return ""
    End Function

    Private Function GetPropertyValue(properties As PropertySet, propertyName As String) As String
        Try
            If properties IsNot Nothing Then
                Return properties.Item(propertyName).Value.ToString()
            End If
        Catch
        End Try
        Return ""
    End Function

    Private Sub SetOrCreateProperty(props As PropertySet, propName As String, valor As Object)
        If props Is Nothing Then Return
        Try
            Dim p As [Property] = Nothing
            Try
                p = props.Item(propName)
            Catch
                p = Nothing
            End Try

            If p Is Nothing Then
                For Each prop As [Property] In props
                    If String.Equals(prop.Name, propName, StringComparison.OrdinalIgnoreCase) Then
                        p = prop
                        Exit For
                    End If
                Next
            End If

            If p IsNot Nothing Then
                p.Value = valor
            Else
                props.Add(valor, propName)
            End If
        Catch
        End Try
    End Sub

    Private Function GetSafePN(d As Document) As String
        Try
            If d IsNot Nothing Then
                Return d.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
            End If
        Catch
        End Try
        Return ""
    End Function

    Private Function GetDocKey(d As Document) As String
        Try
            If d IsNot Nothing Then
                If Not String.IsNullOrEmpty(d.FullDocumentName) Then
                    Return d.FullDocumentName
                Else
                    Return d.DisplayName
                End If
            End If
        Catch
        End Try
        Return Guid.NewGuid().ToString()
    End Function

    Private Function StripMIR(s As String) As String
        If String.IsNullOrEmpty(s) Then Return ""
        Dim t As String = s
        t = t.Replace("_MIR", "")
        t = t.Replace("-MIR", "")
        t = t.Replace(" MIR", "")
        t = t.Replace("MIR", "")
        Return t.Trim()
    End Function

End Class
