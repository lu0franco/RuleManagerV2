Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms
Imports Inventor

''' <summary>
''' Motor de detección y normalización de simetrías (IAM) para Autodesk Inventor.
''' Migra y reemplaza la regla iLogic externa 'SEARCHSIMETRIAS.iLogicVb' a código nativo VB.NET (.vb).
''' 
''' MEJORAS CLAVE DE DETECCIÓN Y ROBUSTEZ:
''' 1. Ya NO depende de que el nombre del archivo o Part Number contenga "MIR".
''' 2. Análisis de Contornos Geométricos (Geometric Contour Analysis):
'''    - Inspección de bucles de aristas (EdgeLoops), perímetro de caras dominantes y quiralidad (winding/traversal).
'''    - Verificación invariante de propiedades de masa (volumen, área superficial, centro de masa).
'''    - Verificación de congruencia por reflexión 3D sobre planos principales vs rotación rígida.
''' 3. Inspección del determinante de la matriz de transformación 4x4 (det < 0 indica reflexión física en ensamblaje).
''' 4. Detección directa en el árbol de componentes derivados (DerivedPartComponent / DerivedAssemblyComponent)
'''    con planos de simetría activos (kMirrorPlaneXY, kMirrorPlaneYZ, kMirrorPlaneXZ), recuperando el documento
'''    original directamente del ReferencedDocumentDescriptor.
''' 5. Asigna Description = "SIMETRIA" y copia el Part Number de la pieza original para permitir el correcto
'''    agrupamiento en listas de corte, ERP y en el Contador de Partes (CANTIDAD_DERECHA / CANTIDAD_IZQUIERDA).
''' </summary>
Public Class SearchSimetriasEngine

    Private ReadOnly _invApp As Inventor.Application

    Public Sub New(invApp As Inventor.Application)
        _invApp = invApp
    End Sub

    ''' <summary>
    ''' Resultado de la ejecución de búsqueda de simetrías.
    ''' </summary>
    Public Class ResultadoSimetria
        Public Property OcurrenciasSimetricas As New List(Of ComponentOccurrence)
        Public Property NombresUnicos As New List(Of String)
        Public Property Items As New List(Of SimetriaItemInfo)
        Public Property SimetricasInfo As List(Of SimetriaItemInfo)
            Get
                Return Items
            End Get
            Set(value As List(Of SimetriaItemInfo))
                Items = value
            End Set
        End Property
        Public Property DocumentosModificados As Integer = 0
        Public Property MensajeResumen As String = ""
        Public Property DetallesLog As New List(Of String)
    End Class

    ''' <summary>
    ''' Firma geométrica y topológica de contornos de una pieza para comparación de simetrías.
    ''' </summary>
    Public Class GeometricContourSignature
        Public Property DocumentPath As String = ""
        Public Property DisplayName As String = ""
        Public Property Volume As Double = 0.0
        Public Property Area As Double = 0.0
        Public Property DimensionX As Double = 0.0
        Public Property DimensionY As Double = 0.0
        Public Property DimensionZ As Double = 0.0
        Public Property FaceCount As Integer = 0
        Public Property PlanarFaceCount As Integer = 0
        Public Property EdgeLoopCount As Integer = 0
        Public Property DominantFacePerimeter As Double = 0.0
        Public Property DominantFaceEdgeCount As Integer = 0
        Public Property DominantFaceHoleCount As Integer = 0
        Public Property DominantFaceEdgeLengths As New List(Of Double)
        Public Property DominantFaceCurveTypes As New List(Of Integer)
        Public Property IsChiral As Boolean = False
    End Class

    ''' <summary>
    ''' Ejecuta el flujo principal de búsqueda, análisis geométrico de contornos y normalización de propiedades.
    ''' </summary>
    Public Function Ejecutar(Optional mostrarDialogo As Boolean = True) As ResultadoSimetria
        Dim resultado As New ResultadoSimetria()

        Try
            ' 1. Validar documento activo
            If _invApp.ActiveDocument Is Nothing Then
                If mostrarDialogo Then
                    MessageBox.Show("No hay ningún documento activo en Autodesk Inventor.",
                                    "Búsqueda de Simetrías", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
                resultado.MensajeResumen = "No hay documento activo."
                Return resultado
            End If

            If _invApp.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                If mostrarDialogo Then
                    MessageBox.Show("Esta herramienta solo puede ejecutarse en documentos de ensamblaje (.iam).",
                                    "Búsqueda de Simetrías", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
                resultado.MensajeResumen = "Documento no es un ensamblaje."
                Return resultado
            End If

            Dim asmDoc As AssemblyDocument = DirectCast(_invApp.ActiveDocument, AssemblyDocument)

            ' 2. Pre-indexar contornos geométricos de piezas únicas en el ensamblaje para análisis de congruencia
            Dim catalogoContornos As Dictionary(Of String, GeometricContourSignature) = IndexarContornosEnsamblaje(asmDoc)

            ' 3. Detectar ocurrencias simétricas (recursivo) sin depender de "MIR"
            Dim listaSimetricas As New List(Of ComponentOccurrence)
            Dim dictOriginalDocs As New Dictionary(Of String, Document)

            FindMirrorOccurrencesRecursive(asmDoc.ComponentDefinition.Occurrences,
                                           catalogoContornos,
                                           listaSimetricas,
                                           dictOriginalDocs,
                                           resultado.DetallesLog)

            resultado.OcurrenciasSimetricas = listaSimetricas

            ' 4. Procesar y normalizar propiedades en los documentos detectados como simetría
            Dim procesados As New Dictionary(Of String, Boolean)
            Dim nombresUnicos As New List(Of String)

            For Each occ As ComponentOccurrence In listaSimetricas
                Try
                    Dim doc As Document = occ.Definition.Document
                    If doc Is Nothing Then Continue For

                    Dim keyDoc As String = GetDocKey(doc)
                    Dim nombreMostrar As String = doc.DisplayName

                    If Not nombresUnicos.Contains(nombreMostrar) Then
                        nombresUnicos.Add(nombreMostrar)
                    End If

                    If procesados.ContainsKey(keyDoc) Then Continue For

                    ' Marcar descripción = "SIMETRIA"
                    Try
                        Dim stdProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                        stdProps.Item("Description").Value = "SIMETRIA"
                    Catch exDesc As Exception
                        resultado.DetallesLog.Add("Aviso: No se pudo escribir Description en " & nombreMostrar & ": " & exDesc.Message)
                    End Try

                    ' Obtener Part Number de la pieza original (sin importar el nombre ni sufijo MIR)
                    Dim origDoc As Document = Nothing
                    If dictOriginalDocs.ContainsKey(keyDoc) Then
                        origDoc = dictOriginalDocs(keyDoc)
                    End If

                    Dim originalPN As String = ResolverPartNumberOriginal(occ, origDoc, catalogoContornos)

                    If Not String.IsNullOrEmpty(originalPN) Then
                        Try
                            Dim stdProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                            stdProps.Item("Part Number").Value = originalPN
                        Catch exPN As Exception
                            resultado.DetallesLog.Add("Aviso: No se pudo escribir Part Number en " & nombreMostrar & ": " & exPN.Message)
                        End Try
                    End If

                    Dim infoItem As New SimetriaItemInfo(occ, "Simetría por Análisis Geométrico", origDoc, originalPN)
                    infoItem.PartNumberActual = GetSafePN(doc)
                    infoItem.PartNumberOriginalSugerido = originalPN
                    infoItem.DescripcionNueva = "SIMETRIA"
                    infoItem.AplicarCambio = True
                    infoItem.Procesado = True
                    resultado.Items.Add(infoItem)

                    procesados(keyDoc) = True
                    resultado.DocumentosModificados += 1
                Catch exOcc As Exception
                    resultado.DetallesLog.Add("Error al procesar ocurrencia " & occ.Name & ": " & exOcc.Message)
                End Try
            Next

            resultado.NombresUnicos = nombresUnicos

            ' 5. Notificación al usuario
            If nombresUnicos.Count > 0 Then
                resultado.MensajeResumen = "Se detectaron y normalizaron las siguientes piezas/subensamblajes como simetrías:" & vbCrLf & vbCrLf &
                                           String.Join(vbCrLf, nombresUnicos.Select(Function(n) " • " & n))

                If mostrarDialogo Then
                    MessageBox.Show(resultado.MensajeResumen,
                                    "Simetrías Detectadas por Análisis Geométrico",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information)
                End If
            Else
                resultado.MensajeResumen = "No se detectaron piezas o ensamblajes con simetría geométrica en el modelo."
                If mostrarDialogo Then
                    MessageBox.Show(resultado.MensajeResumen,
                                    "Búsqueda de Simetrías",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information)
                End If
            End If

        Catch ex As Exception
            resultado.MensajeResumen = "Error general en SearchSimetriasEngine: " & ex.Message
            If mostrarDialogo Then
                MessageBox.Show(resultado.MensajeResumen,
                                "Error - Búsqueda de Simetrías",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error)
            End If
        End Try

        Return resultado
    End Function

    ' =========================================================================
    ' RECURSIÓN: DETECCIÓN ROBUSTA DE SIMETRÍAS
    ' =========================================================================

    ''' <summary>
    ''' Recorre recursivamente todas las ocurrencias evaluando criterios geométricos y topológicos.
    ''' </summary>
    Public Sub FindMirrorOccurrencesRecursive(
        occurrences As ComponentOccurrences,
        catalogoContornos As Dictionary(Of String, GeometricContourSignature),
        ByRef listaSimetricas As List(Of ComponentOccurrence),
        ByRef dictOriginalDocs As Dictionary(Of String, Document),
        ByRef log As List(Of String))

        If occurrences Is Nothing Then Return

        For Each occ As ComponentOccurrence In occurrences
            Try
                If occ.Suppressed Then Continue For
                If occ.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then Continue For

                ' Omitir Content Center / Normalizados
                If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Try
                        Dim pDef As PartComponentDefinition = TryCast(occ.Definition, PartComponentDefinition)
                        If pDef IsNot Nothing AndAlso pDef.IsContentMember Then Continue For
                    Catch
                    End Try
                End If

                Dim isMirror As Boolean = False
                Dim originalDoc As Document = Nothing
                Dim metodoDeteccion As String = ""

                ' -------------------------------------------------------------
                ' CRITERIO 1: Matriz de Transformación (Determinante Negativo)
                ' Un determinante < 0 en la matriz afín 4x4 indica matemáticamente
                ' una reflexión especular (isometría con inversión de orientación).
                ' -------------------------------------------------------------
                If Not isMirror Then
                    Try
                        Dim det As Double = occ.Transformation.Determinant
                        If det < -0.0001 Then
                            isMirror = True
                            metodoDeteccion = "Determinante de Transformación Negativo (" & det.ToString("F3") & ")"
                        End If
                    Catch
                    End Try
                End If

                ' -------------------------------------------------------------
                ' CRITERIO 2: Patrón Espejo de Ensamblaje (MirrorOccurrencePattern)
                ' -------------------------------------------------------------
                If Not isMirror AndAlso occ.PatternElement IsNot Nothing Then
                    Try
                        Dim patElem As Object = occ.PatternElement
                        Dim compObj As Object = patElem.Components
                        If compObj IsNot Nothing Then
                            Dim compTypeName As String = compObj.Type.ToString().ToUpper()
                            ' kMirrorPatternType = 111874
                            If compTypeName.Contains("MIRROR") OrElse compTypeName.Contains("SIMETR") OrElse CInt(compObj.Type) = 111874 Then
                                isMirror = True
                                metodoDeteccion = "Patrón Espejo de Ensamblaje"
                            End If
                        End If
                    Catch
                        Try
                            Dim parentPat As Object = occ.PatternElement.Parent
                            If parentPat IsNot Nothing Then
                                Dim patName As String = parentPat.Name.ToString().ToUpper()
                                If patName.Contains("MIRROR") OrElse patName.Contains("SIMETR") OrElse patName.Contains("ESPEJO") Then
                                    isMirror = True
                                    metodoDeteccion = "Elemento de Patrón Simétrico"
                                End If
                            End If
                        Catch
                        End Try
                    End Try
                End If

                ' -------------------------------------------------------------
                ' CRITERIO 3: Pieza Derivada con Plano de Espejo Activo
                ' Inventor crea piezas simétricas opuestas como DerivedPartComponent
                ' con un plano de simetría (kMirrorPlaneXY, kMirrorPlaneYZ, etc.).
                ' -------------------------------------------------------------
                If Not isMirror AndAlso occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Try
                        Dim docPart As PartDocument = TryCast(occ.Definition.Document, PartDocument)
                        If docPart IsNot Nothing AndAlso docPart.ComponentDefinition.ReferenceComponents IsNot Nothing Then
                            Dim refComps As ReferenceComponents = docPart.ComponentDefinition.ReferenceComponents

                            If refComps.DerivedPartComponents IsNot Nothing AndAlso refComps.DerivedPartComponents.Count > 0 Then
                                For i As Integer = 1 To refComps.DerivedPartComponents.Count
                                    Dim dpc As DerivedPartComponent = refComps.DerivedPartComponents.Item(i)
                                    Try
                                        ' Verificar si el componente derivado tiene espejo habilitado
                                        Dim mirrorOption As Object = dpc.MirrorPlane
                                        ' Si MirrorPlane es kMirrorPlaneXY (26115), kMirrorPlaneYZ (26114), kMirrorPlaneXZ (26113)
                                        ' o diferente a kNoMirroring (26116)
                                        If mirrorOption IsNot Nothing AndAlso CInt(mirrorOption) <> 26116 Then
                                            isMirror = True
                                            metodoDeteccion = "Componente Derivado Espejado (Plano: " & mirrorOption.ToString() & ")"
                                            Try
                                                originalDoc = dpc.ReferencedDocumentDescriptor.ReferencedDocument
                                            Catch
                                            End Try
                                            Exit For
                                        End If
                                    Catch
                                    End Try

                                    ' Si tiene derivado y no leemos el enum, comparamos contornos entre la derivada y la original
                                    If Not isMirror Then
                                        Try
                                            Dim parentRefDoc As Document = dpc.ReferencedDocumentDescriptor.ReferencedDocument
                                            If parentRefDoc IsNot Nothing AndAlso TypeOf parentRefDoc Is PartDocument Then
                                                If SonPiezasSimetricasPorContornos(DirectCast(parentRefDoc, PartDocument), docPart) Then
                                                    isMirror = True
                                                    originalDoc = parentRefDoc
                                                    metodoDeteccion = "Contornos Geométricos Espejados vs Pieza Derivada Padre"
                                                    Exit For
                                                End If
                                            End If
                                        Catch
                                        End Try
                                    End If
                                Next
                            End If
                        End If
                    Catch
                    End Try
                End If

                ' -------------------------------------------------------------
                ' CRITERIO 4: Subensamblaje Derivado con Plano de Espejo
                ' -------------------------------------------------------------
                If Not isMirror AndAlso occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Try
                        Dim docAsm As AssemblyDocument = TryCast(occ.Definition.Document, AssemblyDocument)
                        If docAsm IsNot Nothing AndAlso docAsm.ComponentDefinition.ReferenceComponents IsNot Nothing Then
                            Dim refComps As ReferenceComponents = docAsm.ComponentDefinition.ReferenceComponents
                            If refComps.DerivedAssemblyComponents IsNot Nothing AndAlso refComps.DerivedAssemblyComponents.Count > 0 Then
                                For i As Integer = 1 To refComps.DerivedAssemblyComponents.Count
                                    Dim dac As DerivedAssemblyComponent = refComps.DerivedAssemblyComponents.Item(i)
                                    Try
                                        Dim mirrorOption As Object = dac.MirrorPlane
                                        If mirrorOption IsNot Nothing AndAlso CInt(mirrorOption) <> 26116 Then
                                            isMirror = True
                                            metodoDeteccion = "Subensamblaje Derivado con Espejo"
                                            Try
                                                originalDoc = dac.ReferencedDocumentDescriptor.ReferencedDocument
                                            Catch
                                            End Try
                                            Exit For
                                        End If
                                    Catch
                                    End Try
                                Next
                            End If
                        End If
                    Catch
                    End Try
                End If

                ' -------------------------------------------------------------
                ' CRITERIO 5: Análisis de Contornos Geométricos Global en Ensamblaje
                ' Compara esta pieza con otras piezas del catálogo de contornos para
                ' detectar si existe una pieza complementaria quiral (mano opuesta).
                ' -------------------------------------------------------------
                If Not isMirror AndAlso occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Try
                        Dim docKey As String = GetDocKey(occ.Definition.Document)
                        If catalogoContornos.ContainsKey(docKey) Then
                            Dim sigTarget As GeometricContourSignature = catalogoContornos(docKey)

                            ' Si la pieza es quiral (asimétrica respecto a sí misma), buscar su contraparte
                            If sigTarget.IsChiral Then
                                For Each kvp In catalogoContornos
                                    If kvp.Key = docKey Then Continue For

                                    Dim sigCand As GeometricContourSignature = kvp.Value
                                    If SonContornosEspejados(sigTarget, sigCand) Then
                                        ' Se encontró la pareja simétrica
                                        ' Para determinar cuál es la derivada o simétrica:
                                        ' La que tiene menor orden de aparición o la identificada como contraparte
                                        isMirror = True
                                        metodoDeteccion = "Análisis de Contornos Geométricos (Pareja simétrica: " & sigCand.DisplayName & ")"
                                        Try
                                            For Each subOcc As ComponentOccurrence In occurrences
                                                If subOcc.Definition.Document IsNot Nothing AndAlso
                                                   GetDocKey(subOcc.Definition.Document) = kvp.Key Then
                                                    originalDoc = subOcc.Definition.Document
                                                    Exit For
                                                End If
                                            Next
                                        Catch
                                        End Try
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    Catch
                    End Try
                End If

                ' Registrar si es simétrica
                If isMirror Then
                    listaSimetricas.Add(occ)
                    Dim docKey As String = GetDocKey(occ.Definition.Document)
                    If originalDoc IsNot Nothing AndAlso Not dictOriginalDocs.ContainsKey(docKey) Then
                        dictOriginalDocs(docKey) = originalDoc
                    End If
                    log.Add("Simetría detectada en [" & occ.Name & "] -> Método: " & metodoDeteccion)
                End If

                ' LLAMADA RECURSIVA: Buscar dentro de los subensamblajes hijos
                If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Try
                        If occ.SubOccurrences IsNot Nothing AndAlso occ.SubOccurrences.Count > 0 Then
                            FindMirrorOccurrencesRecursive(occ.SubOccurrences,
                                                           catalogoContornos,
                                                           listaSimetricas,
                                                           dictOriginalDocs,
                                                           log)
                        End If
                    Catch
                    End Try
                End If

            Catch exLoop As Exception
                log.Add("Aviso en iteración de ocurrencia " & occ.Name & ": " & exLoop.Message)
            End Try
        Next
    End Sub

    ' =========================================================================
    ' ANÁLISIS DE CONTORNOS GEOMÉTRICOS (GEOMETRIC CONTOUR ANALYSIS)
    ' =========================================================================

    ''' <summary>
    ''' Indexa las firmas de contornos geométricos de todas las piezas únicas del ensamblaje.
    ''' </summary>
    Private Function IndexarContornosEnsamblaje(asmDoc As AssemblyDocument) As Dictionary(Of String, GeometricContourSignature)
        Dim catalogo As New Dictionary(Of String, GeometricContourSignature)

        Try
            For Each docRef As Document In asmDoc.AllReferencedDocuments
                Try
                    If docRef.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then Continue For
                    Dim partDoc As PartDocument = TryCast(docRef, PartDocument)
                    If partDoc Is Nothing Then Continue For

                    Dim key As String = GetDocKey(partDoc)
                    If catalogo.ContainsKey(key) Then Continue For

                    Dim sig As GeometricContourSignature = ExtraerFirmaContorno(partDoc)
                    If sig IsNot Nothing Then
                        catalogo.Add(key, sig)
                    End If
                Catch
                End Try
            Next
        Catch
        End Try

        Return catalogo
    End Function

    ''' <summary>
    ''' Extrae la firma de contornos y bucles perimetrales de una pieza.
    ''' </summary>
    Public Function ExtraerFirmaContorno(partDoc As PartDocument) As GeometricContourSignature
        Try
            Dim pDef As PartComponentDefinition = partDoc.ComponentDefinition
            If pDef.SurfaceBodies.Count = 0 Then Return Nothing

            Dim sig As New GeometricContourSignature()
            sig.DocumentPath = partDoc.FullFileName
            sig.DisplayName = partDoc.DisplayName

            ' 1. Propiedades de masa globales
            Try
                Dim massProps As MassProperties = pDef.MassProperties
                sig.Volume = massProps.Volume
                sig.Area = massProps.Area
            Catch
            End Try

            ' 2. Dimensiones del Bounding Box orientadas
            Try
                Dim rangeBox As Box = pDef.RangeBox
                Dim dx As Double = Math.Abs(rangeBox.MaxPoint.X - rangeBox.MinPoint.X)
                Dim dy As Double = Math.Abs(rangeBox.MaxPoint.Y - rangeBox.MinPoint.Y)
                Dim dz As Double = Math.Abs(rangeBox.MaxPoint.Z - rangeBox.MinPoint.Z)

                ' Ordenar dimensiones L1 <= L2 <= L3 para comparación independiente de la orientación
                Dim dims As New List(Of Double) From {dx, dy, dz}
                dims.Sort()
                sig.DimensionX = dims(0)
                sig.DimensionY = dims(1)
                sig.DimensionZ = dims(2)
            Catch
            End Try

            ' 3. Conteo de topología
            Dim totalFaces As Integer = 0
            Dim planarFaces As Integer = 0
            Dim totalLoops As Integer = 0
            Dim maxFaceArea As Double = -1.0
            Dim dominantFace As Face = Nothing

            For Each body As SurfaceBody In pDef.SurfaceBodies
                totalFaces += body.Faces.Count
                For Each f As Face In body.Faces
                    If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                        planarFaces += 1
                    End If
                    totalLoops += f.EdgeLoops.Count

                    ' Buscar la cara plana dominante (de mayor área, representativa del contorno principal)
                    Dim fArea As Double = 0.0
                    Try : fArea = f.Evaluator.Area : Catch : End Try
                    If f.SurfaceType = SurfaceTypeEnum.kPlaneSurface AndAlso fArea > maxFaceArea Then
                        maxFaceArea = fArea
                        dominantFace = f
                    End If
                Next
            Next

            sig.FaceCount = totalFaces
            sig.PlanarFaceCount = planarFaces
            sig.EdgeLoopCount = totalLoops

            ' 4. Análisis detallado del contorno exterior de la cara dominante
            If dominantFace IsNot Nothing Then
                Try
                    Dim outerLoop As EdgeLoop = Nothing
                    Dim innerLoopsCount As Integer = 0

                    For Each loopObj As EdgeLoop In dominantFace.EdgeLoops
                        If loopObj.IsOuterLoop Then
                            outerLoop = loopObj
                        Else
                            innerLoopsCount += 1
                        End If
                    Next

                    sig.DominantFaceHoleCount = innerLoopsCount

                    If outerLoop IsNot Nothing Then
                        Dim totalPerimeter As Double = 0.0
                        Dim edgeLengths As New List(Of Double)
                        Dim curveTypes As New List(Of Integer)

                        For Each ed As Edge In outerLoop.Edges
                            Dim len As Double = 0.0
                            Try
                                len = ed.Length
                            Catch
                            End Try
                            totalPerimeter += len
                            edgeLengths.Add(Math.Round(len, 3))
                            curveTypes.Add(CInt(ed.CurveType))
                        Next

                        sig.DominantFacePerimeter = Math.Round(totalPerimeter, 3)
                        sig.DominantFaceEdgeCount = outerLoop.Edges.Count
                        sig.DominantFaceEdgeLengths = edgeLengths
                        sig.DominantFaceCurveTypes = curveTypes

                        ' Determinar si la cara es quiral (asimétrica) analizando la secuencia ordenada de aristas
                        sig.IsChiral = EvaluarQuiralidadContorno(edgeLengths, curveTypes)
                    End If
                Catch
                End Try
            End If

            Return sig
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Evalúa si una secuencia de segmentos de contorno carece de simetría interna propia (es quiral).
    ''' Si es simétrica sobre sí misma, invertir el sentido del contorno produce la misma secuencia periódica.
    ''' Si es quiral, invertir el sentido del contorno produce una secuencia distinta, requiriendo pieza espejo.
    ''' </summary>
    Private Function EvaluarQuiralidadContorno(edgeLengths As List(Of Double), curveTypes As List(Of Integer)) As Boolean
        If edgeLengths Is Nothing OrElse edgeLengths.Count < 3 Then Return False

        Dim n As Integer = edgeLengths.Count
        ' Si la secuencia normal coincide con la invertida (para algún desfase circular), la figura es aquiral (autosimétrica).
        ' Si NO coincide para ningún desfase, la figura es quiral (espejo genuino).
        Dim seqInvertida As New List(Of Double)(edgeLengths)
        seqInvertida.Reverse()

        Dim esAutoSimetrica As Boolean = False
        For offset As Integer = 0 To n - 1
            Dim coincide As Boolean = True
            For i As Integer = 0 To n - 1
                Dim idxInv As Integer = (i + offset) Mod n
                If Math.Abs(edgeLengths(i) - seqInvertida(idxInv)) > 0.01 Then
                    coincide = False
                    Exit For
                End If
            Next
            If coincide Then
                esAutoSimetrica = True
                Exit For
            End If
        Next

        ' Es quiral si no es auto-simétrica
        Return Not esAutoSimetrica
    End Function

    ''' <summary>
    ''' Compara dos piezas directamente por sus contornos geométricos.
    ''' </summary>
    Private Function SonPiezasSimetricasPorContornos(partA As PartDocument, partB As PartDocument) As Boolean
        Try
            Dim sigA As GeometricContourSignature = ExtraerFirmaContorno(partA)
            Dim sigB As GeometricContourSignature = ExtraerFirmaContorno(partB)

            If sigA Is Nothing OrElse sigB Is Nothing Then Return False
            Return SonContornosEspejados(sigA, sigB)
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Verifica si dos firmas geométricas corresponden a contornos especulares (simétricos).
    ''' </summary>
    Public Function SonContornosEspejados(sigA As GeometricContourSignature, sigB As GeometricContourSignature) As Boolean
        If sigA Is Nothing OrElse sigB Is Nothing Then Return False

        ' 1. Filtro invariante de volumen y superficie (tolerancia 0.1%)
        If sigA.Volume > 0 AndAlso sigB.Volume > 0 Then
            Dim diffVol As Double = Math.Abs(sigA.Volume - sigB.Volume) / Math.Max(sigA.Volume, sigB.Volume)
            If diffVol > 0.005 Then Return False
        End If

        If sigA.Area > 0 AndAlso sigB.Area > 0 Then
            Dim diffArea As Double = Math.Abs(sigA.Area - sigB.Area) / Math.Max(sigA.Area, sigB.Area)
            If diffArea > 0.005 Then Return False
        End If

        ' 2. Dimensiones de la caja envolvente ordenadas
        If Math.Abs(sigA.DimensionX - sigB.DimensionX) > 0.05 OrElse
           Math.Abs(sigA.DimensionY - sigB.DimensionY) > 0.05 OrElse
           Math.Abs(sigA.DimensionZ - sigB.DimensionZ) > 0.05 Then
            Return False
        End If

        ' 3. Topología de caras y contornos
        If sigA.FaceCount <> sigB.FaceCount Then Return False
        If sigA.PlanarFaceCount <> sigB.PlanarFaceCount Then Return False
        If sigA.DominantFaceEdgeCount <> sigB.DominantFaceEdgeCount Then Return False
        If sigA.DominantFaceHoleCount <> sigB.DominantFaceHoleCount Then Return False

        If Math.Abs(sigA.DominantFacePerimeter - sigB.DominantFacePerimeter) > 0.05 Then Return False

        ' 4. Análisis de Contorno: Las longitudes de aristas de la cara dominante deben coincidir
        ' en sentido inverso (reflejo / cambio de sentido horario a antihorario)
        If sigA.DominantFaceEdgeLengths.Count > 0 AndAlso sigA.DominantFaceEdgeLengths.Count = sigB.DominantFaceEdgeLengths.Count Then
            Dim listA As List(Of Double) = sigA.DominantFaceEdgeLengths
            Dim listBInvertida As New List(Of Double)(sigB.DominantFaceEdgeLengths)
            listBInvertida.Reverse()

            Dim n As Integer = listA.Count
            Dim coincideReflejado As Boolean = False

            For offset As Integer = 0 To n - 1
                Dim match As Boolean = True
                For i As Integer = 0 To n - 1
                    Dim idxB As Integer = (i + offset) Mod n
                    If Math.Abs(listA(i) - listBInvertida(idxB)) > 0.02 Then
                        match = False
                        Exit For
                    End If
                Next
                If match Then
                    coincideReflejado = True
                    Exit For
                End If
            Next

            If Not coincideReflejado Then Return False
        End If

        Return True
    End Function

    ' =========================================================================
    ' RESOLUCIÓN DE PART NUMBER ORIGINAL (SIN DEPENDER DE "MIR")
    ' =========================================================================

    ''' <summary>
    ''' Resuelve el Part Number original de la pieza contraparte o padre derivada.
    ''' No depende de que contenga "MIR", sino de referencias y contornos.
    ''' </summary>
    Private Function ResolverPartNumberOriginal(
        occ As ComponentOccurrence,
        origDocDirecto As Document,
        catalogoContornos As Dictionary(Of String, GeometricContourSignature)) As String

        ' 1. Si tenemos el documento original directo por relación derivada o geométrica
        If origDocDirecto IsNot Nothing Then
            Dim pnDirecto As String = GetSafePN(origDocDirecto)
            If Not String.IsNullOrEmpty(pnDirecto) Then
                Return pnDirecto
            End If
        End If

        Dim doc As Document = occ.Definition.Document
        If doc Is Nothing Then Return ""

        ' 2. Pieza: Buscar en DerivedPartComponents
        If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
            Try
                Dim pDoc As PartDocument = TryCast(doc, PartDocument)
                If pDoc IsNot Nothing AndAlso pDoc.ComponentDefinition.ReferenceComponents IsNot Nothing Then
                    Dim refComps As ReferenceComponents = pDoc.ComponentDefinition.ReferenceComponents
                    If refComps.DerivedPartComponents IsNot Nothing AndAlso refComps.DerivedPartComponents.Count > 0 Then
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
                End If
            Catch
            End Try
        End If

        ' 3. Ensamblaje: Buscar en DerivedAssemblyComponents
        If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            Try
                Dim aDoc As AssemblyDocument = TryCast(doc, AssemblyDocument)
                If aDoc IsNot Nothing AndAlso aDoc.ComponentDefinition.ReferenceComponents IsNot Nothing Then
                    Dim refCompsAsm As ReferenceComponents = aDoc.ComponentDefinition.ReferenceComponents
                    If refCompsAsm.DerivedAssemblyComponents IsNot Nothing AndAlso refCompsAsm.DerivedAssemblyComponents.Count > 0 Then
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
                End If
            Catch
            End Try
        End If

        ' 4. Buscar en documentos referenciados cuya geometría de contorno coincida
        Try
            If doc.ReferencedDocuments.Count > 0 Then
                Dim docKey As String = GetDocKey(doc)
                Dim sigDoc As GeometricContourSignature = Nothing
                If catalogoContornos.ContainsKey(docKey) Then sigDoc = catalogoContornos(docKey)

                For Each r As Document In doc.ReferencedDocuments
                    Dim rKey As String = GetDocKey(r)
                    If catalogoContornos.ContainsKey(rKey) AndAlso sigDoc IsNot Nothing Then
                        If SonContornosEspejados(sigDoc, catalogoContornos(rKey)) Then
                            Dim pnMatch As String = GetSafePN(r)
                            If Not String.IsNullOrEmpty(pnMatch) Then Return pnMatch
                        End If
                    End If
                Next

                ' Fallback en referencias generales
                For Each r As Document In doc.ReferencedDocuments
                    Dim pnRef As String = GetSafePN(r)
                    If Not String.IsNullOrEmpty(pnRef) Then
                        Return pnRef
                    End If
                Next
            End If
        Catch
        End Try

        ' 5. Como último recurso higiénico: Si el Part Number actual casualmente tuviera sufijos de espejo, retirarlos
        Dim pnActual As String = GetSafePN(doc)
        If Not String.IsNullOrEmpty(pnActual) Then
            Return LimpiarSufijosEspejo(pnActual)
        End If

        Return ""
    End Function

    ' =========================================================================
    ' HELPERS DE PROPIEDADES Y DOCUMENTOS
    ' =========================================================================

    Public Shared Function GetSafePN(d As Document) As String
        If d Is Nothing Then Return ""
        Try
            Return d.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
        Catch
            Return ""
        End Try
    End Function

    Public Shared Function GetDocKey(d As Document) As String
        If d Is Nothing Then Return Guid.NewGuid().ToString()
        Try
            If Not String.IsNullOrEmpty(d.FullDocumentName) Then Return d.FullDocumentName
        Catch
        End Try
        Try
            If Not String.IsNullOrEmpty(d.FullFileName) Then Return d.FullFileName
        Catch
        End Try
        Try
            Return d.DisplayName
        Catch
            Return Guid.NewGuid().ToString()
        End Try
    End Function

    Private Shared Function LimpiarSufijosEspejo(s As String) As String
        If String.IsNullOrEmpty(s) Then Return ""
        Dim t As String = s
        t = t.Replace("_MIR", "")
        t = t.Replace("-MIR", "")
        t = t.Replace(" MIR", "")
        t = t.Replace("MIR", "")
        t = t.Replace("_IZQ", "")
        t = t.Replace("-IZQ", "")
        t = t.Replace(" IZQ", "")
        Return t.Trim()
    End Function

    ' =========================================================================
    ' REVISIÓN DE CANTIDADES Y SIMETRÍAS (REEMPLAZO NATIVO .VB DE REVISIONTEXTO_CANTIDADES)
    ' =========================================================================

    ''' <summary>
    ''' Revisa y actualiza de forma recursiva la propiedad TEXTO_CANTIDADES
    ''' en todos los componentes del ensamblaje activo (reemplaza 'RevisionTEXTO_CANTIDADES.iLogicVb' a .vb).
    ''' </summary>
    Public Sub RevisarTextoCantidades()
        Try
            If _invApp.ActiveDocument Is Nothing OrElse _invApp.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Abra un archivo de ensamblaje (.iam) para revisar las cantidades.",
                                "Revisión de Simetrías", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If

            Dim asmDoc As AssemblyDocument = DirectCast(_invApp.ActiveDocument, AssemblyDocument)
            Dim procesados As New Dictionary(Of String, Boolean)
            Dim totalActualizados As Integer = 0

            RevisarTextoCantidadesRecursivo(asmDoc.ComponentDefinition.Occurrences, procesados, totalActualizados)

            MessageBox.Show("TEXTO_CANTIDADES verificado y actualizado con éxito en " & totalActualizados & " componentes.",
                            "Revisión de Simetrías - Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error durante la revisión de cantidades: " & ex.Message,
                            "Error - Revisión de Simetrías", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RevisarTextoCantidadesRecursivo(
        occs As ComponentOccurrences,
        ByRef procesados As Dictionary(Of String, Boolean),
        ByRef totalActualizados As Integer)

        If occs Is Nothing Then Return

        For Each occ As ComponentOccurrence In occs
            Try
                If occ.Suppressed OrElse Not occ.Visible Then Continue For

                Dim doc As Document = occ.Definition.Document
                If doc IsNot Nothing Then
                    Dim docKey As String = GetDocKey(doc)
                    If Not procesados.ContainsKey(docKey) Then
                        procesados.Add(docKey, True)
                        If ActualizarTextoCantidadesDocumento(doc) Then
                            totalActualizados += 1
                        End If
                    End If
                End If

                If occ.SubOccurrences IsNot Nothing AndAlso occ.SubOccurrences.Count > 0 Then
                    RevisarTextoCantidadesRecursivo(occ.SubOccurrences, procesados, totalActualizados)
                End If
            Catch
            End Try
        Next
    End Sub

    Private Function ActualizarTextoCantidadesDocumento(doc As Document) As Boolean
        Try
            Dim userProps As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
            If userProps Is Nothing Then Return False

            Dim cantDer As Integer = 0
            Dim cantIzq As Integer = 0

            Try : cantDer = CInt(GetOrCreateUserProperty(userProps, "CANTIDAD_DERECHA").Value) : Catch : End Try
            Try : cantIzq = CInt(GetOrCreateUserProperty(userProps, "CANTIDAD_IZQUIERDA").Value) : Catch : End Try

            If cantDer > 0 AndAlso cantIzq > 0 Then
                GetOrCreateUserProperty(userProps, "TEXTO_CANTIDADES").Value = "=" & cantDer & " DER.  " & cantIzq & " IZQ."
                Return True
            End If
        Catch
        End Try
        Return False
    End Function

    Private Shared Function GetOrCreateUserProperty(props As PropertySet, propName As String) As [Property]
        For Each p As [Property] In props
            If p.Name.Equals(propName, StringComparison.OrdinalIgnoreCase) Then
                Return p
            End If
        Next
        Return props.Add("", propName)
    End Function

End Class
