Imports Inventor
Imports System.IO
Imports System.Linq
Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Imports IOPath = System.IO.Path
Imports IOFile = System.IO.File
Imports IODirectory = System.IO.Directory

Public Class RenombradorComponentes

    Private ReadOnly oApp As Inventor.Application
    Private iProcessed As Integer = 0
    Private iSkipped As Integer = 0
    Private iErrors As Integer = 0
    Private sLog As String = ""
    Private bDryRun As Boolean = False
    Private sLogFilePath As String = "C:\Temp\RenombradorLog.txt"

    Private _renamedFiles As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    Private _modelStates As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

    Private _filesToDelete As New List(Of String)()
    Private _hasRoutedSystems As Boolean = False
    Private _routedSystemPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Private _protectedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    Public Event ProgresoReportado(mensaje As String)
    Public Event ProcesoFinalizado(procesados As Integer, omitidos As Integer, errores As Integer, log As String)

    Public Sub New(ByVal invApp As Inventor.Application)
        oApp = invApp
    End Sub

    Public Sub Ejecutar(ByVal sAsmOriginalPath As String, ByVal sPrefix As String, ByVal dryRun As Boolean)
        bDryRun = dryRun
        iProcessed = 0
        iSkipped = 0
        iErrors = 0
        sLog = ""
        _renamedFiles.Clear()
        _modelStates.Clear()
        _filesToDelete.Clear()
        _hasRoutedSystems = False
        _routedSystemPaths.Clear()
        _protectedPaths.Clear()

        sLog = "=== LOG RENOMBRADOR (TOP-DOWN RECURSIVO v9 SilentOp) " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " ===" & vbCrLf
        sLog &= "Ensamblaje raíz: " & sAsmOriginalPath & vbCrLf
        sLog &= "Prefijo: " & sPrefix & vbCrLf
        sLog &= "DryRun: " & bDryRun.ToString() & vbCrLf & vbCrLf

        ' ═══════════════════════════════════════════════════════
        ' FIX: SILENT OPERATION - evita diálogos interactivos
        ' ═══════════════════════════════════════════════════════
        Dim bSilentOriginal As Boolean = False
        Try
            bSilentOriginal = oApp.SilentOperation
            oApp.SilentOperation = True
            RegistrarLog("🔇 SilentOperation activado")
        Catch ex As Exception
            RegistrarLog("⚠️ No se pudo activar SilentOperation: " & ex.Message)
        End Try

        Try
            Dim oAsmDoc As AssemblyDocument = CType(oApp.ActiveDocument, AssemblyDocument)

            Dim oAllOccurrences As New List(Of ComponentOccurrence)()
            ObtenerOcurrenciasRecursivas(oAsmDoc.ComponentDefinition.Occurrences, oAllOccurrences)
            RegistrarLog("Total ocurrencias: " & oAllOccurrences.Count)

            ' === FASE 0: Detectar y marcar todos los routed systems y sus hijos ===
            DetectarRoutedSystemsRecursivo(oAsmDoc.ComponentDefinition.Occurrences, Nothing)
            RegistrarLog("Routed Systems detectados: " & _routedSystemPaths.Count)
            For Each rsPath As String In _routedSystemPaths
                RegistrarLog("   🚫 Routed System protegido: " & IOPath.GetFileName(rsPath))
            Next

            ExtraerTodosLosEstados(oAllOccurrences)
            RegistrarLog("Estados extraídos: " & _modelStates.Count & vbCrLf)

            Dim docToOccurrences As New Dictionary(Of String, List(Of ComponentOccurrence))(StringComparer.OrdinalIgnoreCase)

            For Each oOcc As ComponentOccurrence In oAllOccurrences
                If oOcc Is Nothing Then Continue For

                ' === SKIP TOTAL si es parte de un routed system ===
                If EsParteDeRoutedSystem(oOcc) Then
                    iSkipped += 1
                    Continue For
                End If

                If IsSpecialComponent(oOcc) Then
                    iSkipped += 1
                    Continue For
                End If

                Dim oDoc As Document = Nothing
                If Not TryGetDocumentFromOcc(oOcc, oDoc) OrElse oDoc Is Nothing Then
                    iSkipped += 1
                    Continue For
                End If

                Dim sOriginalPath As String = oDoc.FullFileName
                If String.IsNullOrEmpty(sOriginalPath) OrElse Not IOFile.Exists(sOriginalPath) Then
                    Continue For
                End If

                If Not docToOccurrences.ContainsKey(sOriginalPath) Then
                    docToOccurrences.Add(sOriginalPath, New List(Of ComponentOccurrence)())
                End If
                docToOccurrences(sOriginalPath).Add(oOcc)

                If Not _renamedFiles.ContainsKey(sOriginalPath) Then

                    If EsRoutedSystem(oDoc) Then
                        RegistrarLog("   ⚠️ SKIPPED (Routed System): " & IOPath.GetFileName(sOriginalPath))
                        _renamedFiles.Add(sOriginalPath, sOriginalPath)
                        _hasRoutedSystems = True
                        _protectedPaths.Add(sOriginalPath)
                        iSkipped += 1
                        Continue For
                    End If

                    Dim sBaseNewName As String = BuildNewName(oDoc, sPrefix)
                    If String.IsNullOrEmpty(sBaseNewName) Then
                        iSkipped += 1
                        _renamedFiles.Add(sOriginalPath, sOriginalPath)
                        Continue For
                    End If

                    Dim sFolder As String = IOPath.GetDirectoryName(sOriginalPath)
                    Dim sExt As String = IOPath.GetExtension(sOriginalPath)
                    Dim sNewFullPath As String = IOPath.Combine(sFolder, sBaseNewName & sExt)

                    If sNewFullPath.Length > 250 Then
                        RegistrarLog("   ⚠️ SKIPPED (Path demasiado largo): " & sBaseNewName)
                        _renamedFiles.Add(sOriginalPath, sOriginalPath)
                        iSkipped += 1
                        Continue For
                    End If

                    If IOFile.Exists(sNewFullPath) AndAlso Not String.Equals(sOriginalPath, sNewFullPath, StringComparison.OrdinalIgnoreCase) Then
                        _renamedFiles.Add(sOriginalPath, sNewFullPath)
                        Continue For
                    End If

                    If String.Equals(sOriginalPath, sNewFullPath, StringComparison.OrdinalIgnoreCase) Then
                        _renamedFiles.Add(sOriginalPath, sNewFullPath)
                        iSkipped += 1
                        Continue For
                    End If

                    RegistrarLog("💾 CLONANDO: " & IOPath.GetFileName(sOriginalPath) & " -> " & sBaseNewName & sExt)

                    If Not bDryRun Then
                        If Not ClonarDocumentoRobusto(oDoc, sNewFullPath) Then
                            RegistrarLog("❌ FALLÓ CLONADO de " & IOPath.GetFileName(sOriginalPath) & ", usando original")
                            _renamedFiles.Add(sOriginalPath, sOriginalPath)
                            iErrors += 1
                            Continue For
                        End If
                        ' FIX: Cerrar documento original para evitar diálogo "propiedades actualizadas"
                        CerrarDocumentoSiEstaAbierto(sOriginalPath)
                    End If

                    _renamedFiles.Add(sOriginalPath, sNewFullPath)
                    If Not bDryRun AndAlso Not String.Equals(sOriginalPath, sAsmOriginalPath, StringComparison.OrdinalIgnoreCase) Then
                        _filesToDelete.Add(sOriginalPath)
                    End If
                    iProcessed += 1
                End If
            Next

            RegistrarLog(vbCrLf & "=== FASE REEMPLAZO RECURSIVO (por referencia) ===" & vbCrLf)

            If Not bDryRun Then
                ReemplazarTodasLasOcurrencias(oAllOccurrences)
                oAsmDoc.Update2(True)
            End If

            If Not bDryRun Then
                RegistrarLog(vbCrLf & "=== FASE REEMPLAZO DE REFERENCIAS DE SIMETRÍA Y COMPONENTES DERIVADOS ===" & vbCrLf)
                ActualizarReferenciasDeSimetriaYDerivadas(oAllOccurrences, oAsmDoc)
                oAsmDoc.Update2(True)
            End If

            If Not bDryRun Then
                Try
                    RegistrarLog(vbCrLf & "=== RESTAURANDO ESTADOS ===" & vbCrLf)
                    RestaurarTodosLosEstados(oAllOccurrences)
                Catch ex As Exception
                    RegistrarLog("⚠️ Warning en restauración de estados: " & ex.Message)
                End Try
            End If

            ProcessMainAssembly(oAsmDoc, sPrefix, sAsmOriginalPath)

            Dim iDeleted As Integer = 0
            If Not bDryRun AndAlso iErrors = 0 Then
                RegistrarLog(vbCrLf & "=== ELIMINANDO ARCHIVOS ORIGINALES ===" & vbCrLf)
                iDeleted = EliminarArchivosOriginales()
            ElseIf iErrors > 0 Then
                RegistrarLog(vbCrLf & "⚠️ NO SE ELIMINARON ORIGINALES POR ERRORES PREVIOS" & vbCrLf)
            End If

            MostrarDialogoFinalizacion(iDeleted)

        Catch ex As Exception
            RegistrarLog("❌ ERROR CRÍTICO: " & ex.Message)
            MostrarDialogoFinalizacion(0)
        Finally
            ' ═══════════════════════════════════════════════════════
            ' FIX: RESTAURAR SilentOperation SIEMPRE
            ' ═══════════════════════════════════════════════════════
            Try
                oApp.SilentOperation = bSilentOriginal
                RegistrarLog("🔊 SilentOperation restaurado")
            Catch
            End Try

            Try : IOFile.WriteAllText(sLogFilePath, sLog) : Catch : End Try
            RaiseEvent ProcesoFinalizado(iProcessed, iSkipped, iErrors, sLog)
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════════
    ' FIX: CIERRA DOCUMENTOS ABIERTOS PARA EVITAR DIÁLOGO "PROPIEDADES ACTUALIZADAS"
    ' ═══════════════════════════════════════════════════════════════════════════════

    Private Sub CerrarDocumentoSiEstaAbierto(ByVal sFilePath As String)
        Try
            For Each oDoc As Document In oApp.Documents
                If String.Equals(oDoc.FullFileName, sFilePath, StringComparison.OrdinalIgnoreCase) Then
                    Try
                        ' FIX: Forzar update antes de guardar para evitar cambios pendientes
                        Try : oDoc.Update2(True) : Catch : End Try
                        oDoc.Save2(False)
                        ' Cerrar sin guardar (ya guardamos arriba)
                        oDoc.Close(False)
                        RegistrarLog("   📁 Cerrado para evitar conflicto: " & IOPath.GetFileName(sFilePath))
                    Catch ex As Exception
                        RegistrarLog("   ⚠️ No se pudo cerrar: " & IOPath.GetFileName(sFilePath) & " - " & ex.Message)
                    End Try
                    Exit For
                End If
            Next
        Catch
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════════
    ' DETECCIÓN DE ROUTED SYSTEMS
    ' ═══════════════════════════════════════════════════════════════════════════════

    Private Sub DetectarRoutedSystemsRecursivo(ByVal occs As ComponentOccurrences, ByVal parentRoutedPath As String)
        If occs Is Nothing Then Return

        For Each oOcc As ComponentOccurrence In occs
            If oOcc Is Nothing Then Continue For

            Dim oDoc As Document = Nothing
            If Not TryGetDocumentFromOcc(oOcc, oDoc) OrElse oDoc Is Nothing Then
                Continue For
            End If

            Dim sPath As String = oDoc.FullFileName

            ' Si el padre ya es un routed system, este hijo también lo es
            If Not String.IsNullOrEmpty(parentRoutedPath) Then
                _routedSystemPaths.Add(sPath)
                If oOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Dim subOccs As ComponentOccurrences = Nothing
                    Try : subOccs = oOcc.SubOccurrences : Catch : subOccs = Nothing : End Try
                    If subOccs IsNot Nothing Then
                        DetectarRoutedSystemsRecursivo(subOccs, parentRoutedPath)
                    End If
                End If
                Continue For
            End If

            ' Si este documento ES un routed system, marcarlo y a sus hijos
            If EsRoutedSystem(oDoc) Then
                _routedSystemPaths.Add(sPath)
                _hasRoutedSystems = True
                RegistrarLog("   🔒 Routed System detectado: " & IOPath.GetFileName(sPath))

                If oOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Dim subOccs As ComponentOccurrences = Nothing
                    Try : subOccs = oOcc.SubOccurrences : Catch : subOccs = Nothing : End Try
                    If subOccs IsNot Nothing Then
                        DetectarRoutedSystemsRecursivo(subOccs, sPath)
                    End If
                End If
            Else
                ' No es routed system, pero seguir buscando en sus hijos
                If oOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Dim subOccs As ComponentOccurrences = Nothing
                    Try : subOccs = oOcc.SubOccurrences : Catch : subOccs = Nothing : End Try
                    If subOccs IsNot Nothing Then
                        DetectarRoutedSystemsRecursivo(subOccs, Nothing)
                    End If
                End If
            End If
        Next
    End Sub

    Private Function EsParteDeRoutedSystem(ByVal oOcc As ComponentOccurrence) As Boolean
        If oOcc Is Nothing Then Return False

        Dim oDoc As Document = Nothing
        If TryGetDocumentFromOcc(oOcc, oDoc) AndAlso oDoc IsNot Nothing Then
            If _routedSystemPaths.Contains(oDoc.FullFileName) Then
                Return True
            End If
        End If

        ' Verificar ancestros
        Dim parent As ComponentOccurrence = Nothing
        Try
            parent = oOcc.ParentOccurrence
        Catch
            Return False
        End Try

        While parent IsNot Nothing
            Dim parentDoc As Document = Nothing
            If TryGetDocumentFromOcc(parent, parentDoc) AndAlso parentDoc IsNot Nothing Then
                If _routedSystemPaths.Contains(parentDoc.FullFileName) Then
                    Return True
                End If
            End If
            Try
                parent = parent.ParentOccurrence
            Catch
                Exit While
            End Try
        End While

        Return False
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════════
    ' REEMPLAZO POR REFERENCIA DE ARCHIVO (preserva constraints)
    ' Procesa de abajo hacia arriba: primero los subensamblajes clonados,
    ' luego el padre, usando FileDescriptor.ReplaceReference en cada .iam
    ' ═══════════════════════════════════════════════════════════════════════════════

    Private Sub ReemplazarTodasLasOcurrencias(ByVal oAllOccurrences As List(Of ComponentOccurrence))
        If bDryRun Then Return

        ' 1. Armar un mapa: Documento Padre -> Lista de hijos originales a reemplazar
        Dim parentToChildren As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)

        For Each oOcc As ComponentOccurrence In oAllOccurrences
            If oOcc Is Nothing Then Continue For
            If EsParteDeRoutedSystem(oOcc) Then Continue For
            If IsSpecialComponent(oOcc) Then Continue For

            Dim oDoc As Document = Nothing
            If Not TryGetDocumentFromOcc(oOcc, oDoc) OrElse oDoc Is Nothing Then Continue For

            Dim sOriginalPath As String = oDoc.FullFileName
            If Not _renamedFiles.ContainsKey(sOriginalPath) Then Continue For

            Dim sNewPath As String = _renamedFiles(sOriginalPath)
            If String.Equals(sOriginalPath, sNewPath, StringComparison.OrdinalIgnoreCase) Then Continue For

            ' ¿Quién es el padre de esta ocurrencia?
            Dim oParentDoc As Document = Nothing
            Try
                Dim oParentOcc As ComponentOccurrence = oOcc.ParentOccurrence
                If oParentOcc IsNot Nothing Then
                    Dim oParentDef As ComponentDefinition = oParentOcc.Definition
                    If TypeOf oParentDef Is AssemblyComponentDefinition Then
                        oParentDoc = CType(oParentDef, AssemblyComponentDefinition).Document
                    End If
                Else
                    ' Ocurrencia directa en el raíz
                    oParentDoc = oApp.ActiveDocument
                End If
            Catch
            End Try

            If oParentDoc Is Nothing OrElse oParentDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then Continue For

            Dim sParentPath As String = oParentDoc.FullFileName

            ' Si el padre también fue renombrado, la actualización debe hacerse en el PADRE CLONADO
            If _renamedFiles.ContainsKey(sParentPath) Then
                Dim sNewParentPath As String = _renamedFiles(sParentPath)
                If Not String.Equals(sParentPath, sNewParentPath, StringComparison.OrdinalIgnoreCase) Then
                    sParentPath = sNewParentPath
                End If
            End If

            If Not parentToChildren.ContainsKey(sParentPath) Then
                parentToChildren.Add(sParentPath, New List(Of String))
            End If
            If Not parentToChildren(sParentPath).Contains(sOriginalPath) Then
                parentToChildren(sParentPath).Add(sOriginalPath)
            End If
        Next

        ' 2. Ordenar padres de profundidad máxima a mínima (bottom-up)
        Dim orderedParents = parentToChildren.Keys _
            .OrderByDescending(Function(p) GetProfundidadDeDocumento(p, oAllOccurrences)) _
            .ToList()

        ' 3. Procesar cada padre
        For Each sParentPath In orderedParents
            If _routedSystemPaths.Contains(sParentPath) Then Continue For
            ReemplazarReferenciasEnDocumento(sParentPath, parentToChildren(sParentPath))
        Next
    End Sub

    ' Helper: reemplaza las referencias de archivo dentro de un único .iam
    Private Sub ReemplazarReferenciasEnDocumento(ByVal sDocPath As String, ByVal childrenToReplace As List(Of String))
        If String.IsNullOrEmpty(sDocPath) OrElse Not IOFile.Exists(sDocPath) Then Return
        If childrenToReplace Is Nothing OrElse childrenToReplace.Count = 0 Then Return

        Dim oDoc As Document = Nothing
        Dim bWeOpenedIt As Boolean = False

        Try
            ' Buscar si ya está abierto en la sesión
            For Each d As Document In oApp.Documents
                If String.Equals(d.FullFileName, sDocPath, StringComparison.OrdinalIgnoreCase) Then
                    oDoc = d
                    Exit For
                End If
            Next

            ' Si no está abierto, abrirlo silenciosamente
            If oDoc Is Nothing Then
                oDoc = oApp.Documents.Open(sDocPath, False)
                bWeOpenedIt = True
            End If

            If Not (TypeOf oDoc Is AssemblyDocument) Then
                If bWeOpenedIt Then
                    Try : oDoc.Close(False) : Catch : End Try
                End If
                Return
            End If

            Dim oAsm As AssemblyDocument = CType(oDoc, AssemblyDocument)
            Dim oFile As Inventor.File = oAsm.File
            Dim iCount As Integer = oFile.ReferencedFileDescriptors.Count
            Dim bModified As Boolean = False

            For i As Integer = 1 To iCount
                Dim oFD As FileDescriptor = Nothing
                Try
                    oFD = oFile.ReferencedFileDescriptors.Item(i)
                Catch
                    Continue For
                End Try
                If oFD Is Nothing Then Continue For

                Dim sRefPath As String = ""
                Try
                    sRefPath = oFD.FullFileName
                Catch
                    Continue For
                End Try

                If String.IsNullOrEmpty(sRefPath) Then Continue For
                If Not childrenToReplace.Contains(sRefPath, StringComparer.OrdinalIgnoreCase) Then Continue For
                If Not _renamedFiles.ContainsKey(sRefPath) Then Continue For

                Dim sNewPath As String = _renamedFiles(sRefPath)
                If String.Equals(sRefPath, sNewPath, StringComparison.OrdinalIgnoreCase) Then Continue For
                If Not IOFile.Exists(sNewPath) Then
                    RegistrarLog("  └─ ⚠️ Clon no existe, se omite: " & IOPath.GetFileName(sNewPath))
                    Continue For
                End If

                Try
                    oFD.ReplaceReference(sNewPath)
                    RegistrarLog("🔄 " & IOPath.GetFileName(sRefPath) & " -> " & IOPath.GetFileName(sNewPath) & "  (en " & IOPath.GetFileName(sDocPath) & ")")
                    bModified = True
                Catch ex As Exception
                    RegistrarLog("❌ ERROR reemplazando ref [" & IOPath.GetFileName(sRefPath) & "] en " & IOPath.GetFileName(sDocPath) & ": " & ex.Message)
                End Try
            Next

            If bModified Then
                oAsm.Update2(True)
                oAsm.Save2(False)
                RegistrarLog("💾 Guardado: " & IOPath.GetFileName(sDocPath))
            End If

            ' FIX: Cerrar si lo abrimos nosotros, SIEMPRE (incluso si es activo, con guardar)
            If bWeOpenedIt AndAlso oDoc IsNot Nothing Then
                Try
                    oDoc.Close(True)
                Catch
                End Try
            End If

        Catch ex As Exception
            RegistrarLog("❌ ERROR procesando documento " & IOPath.GetFileName(sDocPath) & ": " & ex.Message)
            If oDoc IsNot Nothing AndAlso bWeOpenedIt Then
                Try
                    oDoc.Close(False)
                Catch
                End Try
            End If
        End Try
    End Sub

    ' Devuelve la profundidad máxima de cualquier ocurrencia que use este documento
    Private Function GetProfundidadDeDocumento(ByVal sDocPath As String, ByVal oAllOccurrences As List(Of ComponentOccurrence)) As Integer
        Dim maxDepth As Integer = 0
        For Each oOcc As ComponentOccurrence In oAllOccurrences
            Try
                Dim oDoc As Document = Nothing
                If TryGetDocumentFromOcc(oOcc, oDoc) AndAlso oDoc IsNot Nothing Then
                    If String.Equals(oDoc.FullFileName, sDocPath, StringComparison.OrdinalIgnoreCase) Then
                        Dim d As Integer = GetNivelProfundidad(oOcc)
                        If d > maxDepth Then maxDepth = d
                    End If
                End If
            Catch
            End Try
        Next
        Return maxDepth
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════════
    ' ACTUALIZACIÓN DE REFERENCIAS EN PIEZAS CON SIMETRÍA Y COMPONENTES DERIVADOS
    ' ═══════════════════════════════════════════════════════════════════════════════

    Private Sub ActualizarReferenciasDeSimetriaYDerivadas(ByVal oAllOccurrences As List(Of ComponentOccurrence), ByVal oAsmDoc As AssemblyDocument)
        RegistrarLog("🔍 Buscando piezas con simetría y componentes derivados para reasignar referencias a los nuevos nombres...")

        Dim filesToInspect As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        ' 1. Incluir todos los archivos nuevos/clonados
        For Each kvp In _renamedFiles
            If Not String.IsNullOrEmpty(kvp.Value) AndAlso IOFile.Exists(kvp.Value) Then
                filesToInspect.Add(kvp.Value)
            End If
        Next

        ' 2. Incluir todos los documentos de las ocurrencias actuales (por si alguno con simetría no fue clonado/renombrado)
        For Each oOcc As ComponentOccurrence In oAllOccurrences
            If oOcc Is Nothing Then Continue For
            Try
                Dim d As Document = Nothing
                If TryGetDocumentFromOcc(oOcc, d) AndAlso d IsNot Nothing Then
                    Dim p As String = d.FullFileName
                    If Not String.IsNullOrEmpty(p) AndAlso IOFile.Exists(p) Then
                        If _renamedFiles.ContainsKey(p) AndAlso IOFile.Exists(_renamedFiles(p)) Then
                            filesToInspect.Add(_renamedFiles(p))
                        Else
                            filesToInspect.Add(p)
                        End If
                    End If
                End If
            Catch
            End Try
        Next

        ' 3. Incluir el ensamblaje raíz
        If oAsmDoc IsNot Nothing AndAlso Not String.IsNullOrEmpty(oAsmDoc.FullFileName) Then
            Dim rootPath As String = oAsmDoc.FullFileName
            If _renamedFiles.ContainsKey(rootPath) AndAlso IOFile.Exists(_renamedFiles(rootPath)) Then
                filesToInspect.Add(_renamedFiles(rootPath))
            Else
                filesToInspect.Add(rootPath)
            End If
        End If

        RegistrarLog("   📁 Total de archivos a examinar para referencias de simetría/derivadas: " & filesToInspect.Count)

        Dim countActualizados As Integer = 0

        For Each sFilePath As String In filesToInspect
            If _routedSystemPaths.Contains(sFilePath) Then Continue For

            If ProcesarReferenciasDeDocumentoIndividual(sFilePath) Then
                countActualizados += 1
            End If
        Next

        RegistrarLog("✅ Total de archivos con simetría/derivadas actualizados: " & countActualizados)
    End Sub

    Private Function ProcesarReferenciasDeDocumentoIndividual(ByVal sDocPath As String) As Boolean
        If String.IsNullOrEmpty(sDocPath) OrElse Not IOFile.Exists(sDocPath) Then Return False

        Dim oDoc As Document = Nothing
        Dim bWeOpenedIt As Boolean = False
        Dim bModified As Boolean = False

        Try
            ' 1. Buscar si ya está abierto en la sesión de Inventor
            For Each d As Document In oApp.Documents
                If String.Equals(d.FullFileName, sDocPath, StringComparison.OrdinalIgnoreCase) Then
                    oDoc = d
                    Exit For
                End If
            Next

            ' 2. Si no está abierto, abrirlo de forma invisible
            If oDoc Is Nothing Then
                oDoc = oApp.Documents.Open(sDocPath, False)
                bWeOpenedIt = True
            End If

            If oDoc Is Nothing Then Return False

            ' === ESTRATEGIA 1: Componentes derivados en piezas (PartDocument -> ReferenceComponents) ===
            If TypeOf oDoc Is PartDocument Then
                Dim partDoc As PartDocument = CType(oDoc, PartDocument)
                Dim partDef As PartComponentDefinition = partDoc.ComponentDefinition

                If partDef.ReferenceComponents IsNot Nothing Then
                    ' A. DerivedPartComponents (Simetría de piezas / Opposite Hand / Piezas derivadas)
                    Try
                        If partDef.ReferenceComponents.DerivedPartComponents IsNot Nothing Then
                            For Each oDP As DerivedPartComponent In partDef.ReferenceComponents.DerivedPartComponents
                                Try
                                    Dim oFD As FileDescriptor = Nothing
                                    Try
                                        If oDP.ReferencedDocumentDescriptor IsNot Nothing Then
                                            oFD = oDP.ReferencedDocumentDescriptor.ReferencedFileDescriptor
                                        End If
                                    Catch
                                    End Try

                                    If oFD Is Nothing Then
                                        Try : oFD = oDP.ReferencedFileDescriptor : Catch : End Try
                                    End If

                                    If oFD IsNot Nothing Then
                                        If IntentarReemplazarFileDescriptor(oFD, sDocPath, "Simetría (DerivedPart)") Then
                                            bModified = True
                                        End If
                                    End If
                                Catch exDP As Exception
                                    RegistrarLog("   ⚠️ Error al revisar DerivedPartComponent en " & IOPath.GetFileName(sDocPath) & ": " & exDP.Message)
                                End Try
                            Next
                        End If
                    Catch
                    End Try

                    ' B. DerivedAssemblyComponents
                    Try
                        If partDef.ReferenceComponents.DerivedAssemblyComponents IsNot Nothing Then
                            For Each oDA As DerivedAssemblyComponent In partDef.ReferenceComponents.DerivedAssemblyComponents
                                Try
                                    Dim oFD As FileDescriptor = Nothing
                                    Try
                                        If oDA.ReferencedDocumentDescriptor IsNot Nothing Then
                                            oFD = oDA.ReferencedDocumentDescriptor.ReferencedFileDescriptor
                                        End If
                                    Catch
                                    End Try

                                    If oFD Is Nothing Then
                                        Try : oFD = oDA.ReferencedFileDescriptor : Catch : End Try
                                    End If

                                    If oFD IsNot Nothing Then
                                        If IntentarReemplazarFileDescriptor(oFD, sDocPath, "Simetría (DerivedAssembly)") Then
                                            bModified = True
                                        End If
                                    End If
                                Catch exDA As Exception
                                    RegistrarLog("   ⚠️ Error al revisar DerivedAssemblyComponent en " & IOPath.GetFileName(sDocPath) & ": " & exDA.Message)
                                End Try
                            Next
                        End If
                    Catch
                    End Try
                End If
            End If

            ' === ESTRATEGIA 2: Componentes derivados en ensamblajes (AssemblyDocument -> ReferenceComponents) ===
            If TypeOf oDoc Is AssemblyDocument Then
                Dim asmDoc As AssemblyDocument = CType(oDoc, AssemblyDocument)
                Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition

                If asmDef.ReferenceComponents IsNot Nothing Then
                    Try
                        If asmDef.ReferenceComponents.DerivedPartComponents IsNot Nothing Then
                            For Each oDP As DerivedPartComponent In asmDef.ReferenceComponents.DerivedPartComponents
                                Try
                                    Dim oFD As FileDescriptor = Nothing
                                    If oDP.ReferencedDocumentDescriptor IsNot Nothing Then
                                        oFD = oDP.ReferencedDocumentDescriptor.ReferencedFileDescriptor
                                    End If
                                    If oFD IsNot Nothing Then
                                        If IntentarReemplazarFileDescriptor(oFD, sDocPath, "Simetría en IAM (DerivedPart)") Then
                                            bModified = True
                                        End If
                                    End If
                                Catch
                                End Try
                            Next
                        End If
                    Catch
                    End Try

                    Try
                        If asmDef.ReferenceComponents.DerivedAssemblyComponents IsNot Nothing Then
                            For Each oDA As DerivedAssemblyComponent In asmDef.ReferenceComponents.DerivedAssemblyComponents
                                Try
                                    Dim oFD As FileDescriptor = Nothing
                                    If oDA.ReferencedDocumentDescriptor IsNot Nothing Then
                                        oFD = oDA.ReferencedDocumentDescriptor.ReferencedFileDescriptor
                                    End If
                                    If oFD IsNot Nothing Then
                                        If IntentarReemplazarFileDescriptor(oFD, sDocPath, "Simetría en IAM (DerivedAssembly)") Then
                                            bModified = True
                                        End If
                                    End If
                                Catch
                                End Try
                            Next
                        End If
                    Catch
                    End Try
                End If
            End If

            ' === ESTRATEGIA 3: ReferencedFileDescriptors generales del documento ===
            Try
                Dim oFile As Inventor.File = oDoc.File
                If oFile IsNot Nothing AndAlso oFile.ReferencedFileDescriptors IsNot Nothing Then
                    For i As Integer = 1 To oFile.ReferencedFileDescriptors.Count
                        Dim oFD As FileDescriptor = Nothing
                        Try : oFD = oFile.ReferencedFileDescriptors.Item(i) : Catch : Continue For : End Try
                        If oFD IsNot Nothing Then
                            If IntentarReemplazarFileDescriptor(oFD, sDocPath, "FileDescriptor") Then
                                bModified = True
                            End If
                        End If
                    Next
                End If
            Catch ex As Exception
            End Try

            ' === ESTRATEGIA 4: ReferencedDocumentDescriptors generales del documento ===
            Try
                If oDoc.ReferencedDocumentDescriptors IsNot Nothing Then
                    For i As Integer = 1 To oDoc.ReferencedDocumentDescriptors.Count
                        Dim oDocDesc As DocumentDescriptor = Nothing
                        Try : oDocDesc = oDoc.ReferencedDocumentDescriptors.Item(i) : Catch : Continue For : End Try
                        If oDocDesc IsNot Nothing AndAlso oDocDesc.ReferencedFileDescriptor IsNot Nothing Then
                            If IntentarReemplazarFileDescriptor(oDocDesc.ReferencedFileDescriptor, sDocPath, "DocumentDescriptor") Then
                                bModified = True
                            End If
                        End If
                    Next
                End If
            Catch ex As Exception
            End Try

            ' Si hubo cambios, forzar actualización y guardado
            If bModified Then
                Try : oDoc.Update2(True) : Catch : Try : oDoc.Update() : Catch : End Try : End Try
                Try : oDoc.Save2(False) : Catch : Try : oDoc.Save() : Catch : End Try : End Try
                RegistrarLog("💾 Guardados cambios de simetría en: " & IOPath.GetFileName(sDocPath))
            End If

            ' Si lo abrimos nosotros, cerrarlo
            If bWeOpenedIt AndAlso oDoc IsNot Nothing Then
                Try
                    oDoc.Close(False)
                Catch
                End Try
            End If

            Return bModified

        Catch ex As Exception
            RegistrarLog("❌ Error procesando simetría en " & IOPath.GetFileName(sDocPath) & ": " & ex.Message)
            If bWeOpenedIt AndAlso oDoc IsNot Nothing Then
                Try : oDoc.Close(False) : Catch : End Try
            End If
            Return False
        End Try
    End Function

    Private Function IntentarReemplazarFileDescriptor(ByVal oFD As FileDescriptor, ByVal sDocOwnerPath As String, ByVal sContexto As String) As Boolean
        If oFD Is Nothing Then Return False

        Dim sRefPath As String = ""
        Try
            sRefPath = oFD.FullFileName
        Catch
        End Try

        If String.IsNullOrEmpty(sRefPath) Then
            Try : sRefPath = oFD.LogicalFileName : Catch : End Try
        End If

        If String.IsNullOrEmpty(sRefPath) Then Return False

        Dim sNewPath As String = ObtenerNuevoPathSiFueRenombrado(sRefPath)
        If String.IsNullOrEmpty(sNewPath) Then Return False
        If String.Equals(sRefPath, sNewPath, StringComparison.OrdinalIgnoreCase) Then Return False

        If Not IOFile.Exists(sNewPath) Then
            RegistrarLog("   ⚠️ Clon no existe para simetría: " & IOPath.GetFileName(sNewPath))
            Return False
        End If

        Try
            oFD.ReplaceReference(sNewPath)
            RegistrarLog("   🔄 [" & sContexto & "] " & IOPath.GetFileName(sRefPath) & " -> " & IOPath.GetFileName(sNewPath) & " (en " & IOPath.GetFileName(sDocOwnerPath) & ")")
            Return True
        Catch ex As Exception
            RegistrarLog("   ❌ ERROR ReplaceReference (" & sContexto & "): " & IOPath.GetFileName(sRefPath) & " -> " & IOPath.GetFileName(sNewPath) & " - " & ex.Message)
            Return False
        End Try
    End Function

    Private Function ObtenerNuevoPathSiFueRenombrado(ByVal sRefPath As String) As String
        If String.IsNullOrEmpty(sRefPath) Then Return Nothing

        ' 1. Búsqueda exacta por ruta completa
        If _renamedFiles.ContainsKey(sRefPath) Then
            Return _renamedFiles(sRefPath)
        End If

        ' 2. Búsqueda por nombre de archivo (por si la referencia tiene ruta relativa o normalizada distinta)
        Dim sFileName As String = IOPath.GetFileName(sRefPath)
        For Each kvp In _renamedFiles
            If String.Equals(IOPath.GetFileName(kvp.Key), sFileName, StringComparison.OrdinalIgnoreCase) Then
                Return kvp.Value
            End If
        Next

        Return Nothing
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════════
    ' HELPERS Y RESTO DE LA CLASE
    ' ═══════════════════════════════════════════════════════════════════════════════

    Private Function ClonarDocumentoRobusto(ByVal oDoc As Document, ByVal sNewPath As String) As Boolean

        Try
            Dim oDocToClone As Document = ObtenerDocumentoParaClonar(oDoc)

            If Not String.Equals(oDocToClone.FullFileName, oDoc.FullFileName, StringComparison.OrdinalIgnoreCase) Then
                Try : oDocToClone.Update2(True) : Catch : End Try
                oDocToClone.SaveAs(sNewPath, True)
                RegistrarLog("   ✅ Clonado desde FactoryDocument")
                Return True
            End If
        Catch ex As Exception
            RegistrarLog("   ⚠️ Estrategia 1 (Factory) falló: " & ex.Message)
        End Try

        Try
            Try : oDoc.Update2(True) : Catch : End Try
            oDoc.SaveAs(sNewPath, True)
            RegistrarLog("   ✅ Clonado desde documento original")
            Return True
        Catch ex As Exception
            RegistrarLog("   ⚠️ Estrategia 2 (Original) falló: " & ex.Message)
        End Try

        Try
            oDoc.SaveAs(sNewPath, False)
            RegistrarLog("   ✅ Clonado con SaveAs(False) — ATENCIÓN: referencias no copiadas")
            Return True
        Catch ex As Exception
            RegistrarLog("   ⚠️ Estrategia 3 (SaveAs False) falló: " & ex.Message)
        End Try

        Try
            System.Threading.Thread.Sleep(1000)
            oDoc.SaveAs(sNewPath, True)
            RegistrarLog("   ✅ Clonado con reintento")
            Return True
        Catch ex As Exception
            RegistrarLog("   ❌ Estrategia 4 (Reintento) falló: " & ex.Message)
        End Try

        Return False
    End Function

    Private Function ObtenerDocumentoParaClonar(ByVal oDoc As Document) As Document
        Try
            If TypeOf oDoc Is PartDocument Then
                Dim oPartDoc As PartDocument = CType(oDoc, PartDocument)
                If oPartDoc.ComponentDefinition.IsModelStateMember Then
                    Return CType(oPartDoc.ComponentDefinition.FactoryDocument, Document)
                End If
            ElseIf TypeOf oDoc Is AssemblyDocument Then
                Dim oAsmDoc As AssemblyDocument = CType(oDoc, AssemblyDocument)
                If oAsmDoc.ComponentDefinition.IsModelStateMember Then
                    Return CType(oAsmDoc.ComponentDefinition.FactoryDocument, Document)
                End If
            End If
        Catch
        End Try
        Return oDoc
    End Function

    Private Sub ExtraerTodosLosEstados(ByVal oAllOccurrences As List(Of ComponentOccurrence))
        For Each oOcc As ComponentOccurrence In oAllOccurrences
            If oOcc Is Nothing Then Continue For

            If EsParteDeRoutedSystem(oOcc) Then
                Continue For
            End If

            Try
                Dim sState As String = ""

                Try
                    sState = oOcc.ActiveModelState
                Catch
                    sState = ""
                End Try

                If String.IsNullOrEmpty(sState) Then
                    Try
                        Dim oDef As ComponentDefinition = oOcc.Definition
                        If TypeOf oDef Is PartComponentDefinition Then
                            Dim oPartDef As PartComponentDefinition = CType(oDef, PartComponentDefinition)
                            If oPartDef.ModelStates.Count > 0 Then
                                sState = oPartDef.ActiveModelState.Name
                            End If
                        ElseIf TypeOf oDef Is AssemblyComponentDefinition Then
                            Dim oAsmDef As AssemblyComponentDefinition = CType(oDef, AssemblyComponentDefinition)
                            If oAsmDef.ModelStates.Count > 0 Then
                                sState = oAsmDef.ActiveModelState.Name
                            End If
                        End If
                    Catch
                        sState = ""
                    End Try
                End If

                If Not String.IsNullOrEmpty(sState) Then
                    Dim sKey As String = GenerarClaveUnica(oOcc)
                    _modelStates(sKey) = sState
                End If

            Catch
            End Try
        Next
    End Sub

    Private Sub RestaurarTodosLosEstados(ByVal oAllOccurrences As List(Of ComponentOccurrence))
        For Each oOcc As ComponentOccurrence In oAllOccurrences
            If oOcc Is Nothing Then Continue For

            Try
                If EsParteDeRoutedSystem(oOcc) Then
                    Continue For
                End If

                If PadreSeraReemplazado(oOcc) Then
                    Continue For
                End If

                Dim bSuppressed As Boolean = False
                Try
                    bSuppressed = oOcc.Suppressed
                Catch
                    Continue For
                End Try
                If bSuppressed Then Continue For

                Dim sKey As String = GenerarClaveUnica(oOcc)
                If Not _modelStates.ContainsKey(sKey) Then Continue For

                Dim sExpectedState As String = _modelStates(sKey)
                If String.IsNullOrEmpty(sExpectedState) Then Continue For

                Dim oDef As ComponentDefinition = Nothing
                Try
                    oDef = oOcc.Definition
                Catch
                    Continue For
                End Try
                If oDef Is Nothing Then Continue For

                Dim oModelStates As ModelStates = Nothing
                Dim sExactName As String = ""
                Dim bExists As Boolean = False

                If TypeOf oDef Is PartComponentDefinition Then
                    oModelStates = CType(oDef, PartComponentDefinition).ModelStates
                ElseIf TypeOf oDef Is AssemblyComponentDefinition Then
                    oModelStates = CType(oDef, AssemblyComponentDefinition).ModelStates
                End If

                If oModelStates Is Nothing Then Continue For

                For Each oState As ModelState In oModelStates
                    If String.Equals(oState.Name, sExpectedState, StringComparison.OrdinalIgnoreCase) Then
                        bExists = True
                        sExactName = oState.Name
                        Exit For
                    End If
                Next

                If Not bExists Then
                    RegistrarLog("⚠️ Estado [" & sExpectedState & "] no existe en clon de " & oOcc.Name)
                    Continue For
                End If

                Dim intentos As Integer = 0
                Dim bAplicado As Boolean = False
                Dim sErrorMsg As String = ""
                Dim sOccName As String = ""

                Try
                    sOccName = oOcc.Name
                Catch
                    sOccName = "ocurrencia_desconocida"
                End Try

                While intentos < 3 AndAlso Not bAplicado
                    intentos += 1
                    Try
                        oOcc.ActiveModelState = sExactName
                        System.Threading.Thread.Sleep(50)

                        Dim sCurrentState As String = ""
                        Try
                            sCurrentState = oOcc.ActiveModelState
                        Catch
                        End Try

                        If String.Equals(sCurrentState, sExactName, StringComparison.OrdinalIgnoreCase) Then
                            bAplicado = True
                            RegistrarLog("✅ Estado [" & sExactName & "] restaurado en " & sOccName)
                        End If
                    Catch ex As Exception
                        sErrorMsg = ex.Message
                        If intentos >= 3 Then
                            RegistrarLog("⚠️ No se pudo restaurar estado [" & sExpectedState & "] en " & sOccName & ": " & sErrorMsg)
                        End If
                        System.Threading.Thread.Sleep(100)
                    End Try
                End While

            Catch ex As Exception
                Dim sOccName As String = "desconocida"
                Try
                    sOccName = oOcc.Name
                Catch
                End Try
                RegistrarLog("⚠️ Warning en restauración de estado para " & sOccName & ": " & ex.Message)
            End Try
        Next
    End Sub

    Private Function PadreSeraReemplazado(ByVal oOcc As ComponentOccurrence) As Boolean
        Dim parent As ComponentOccurrence = Nothing
        Try
            parent = oOcc.ParentOccurrence
        Catch
            Return False
        End Try

        If parent Is Nothing Then Return False

        Dim oParentDoc As Document = Nothing
        If Not TryGetDocumentFromOcc(parent, oParentDoc) OrElse oParentDoc Is Nothing Then Return False

        Dim sParentPath As String = oParentDoc.FullFileName
        If Not _renamedFiles.ContainsKey(sParentPath) Then Return False

        Dim sNewParentPath As String = _renamedFiles(sParentPath)
        Return Not String.Equals(sParentPath, sNewParentPath, StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function GetNivelProfundidad(ByVal oOcc As ComponentOccurrence) As Integer
        Dim nivel As Integer = 0
        Dim current As ComponentOccurrence = oOcc
        Try
            While True
                Dim parent As ComponentOccurrence = Nothing
                Try
                    parent = current.ParentOccurrence
                Catch
                    Exit While
                End Try
                If parent Is Nothing Then Exit While
                nivel += 1
                current = parent
            End While
        Catch
        End Try
        Return nivel
    End Function

    Private Sub ObtenerOcurrenciasRecursivas(ByVal occs As ComponentOccurrences, ByRef lista As List(Of ComponentOccurrence))
        If occs Is Nothing Then Return

        For Each oOcc As ComponentOccurrence In occs
            If oOcc Is Nothing Then Continue For

            Try
                lista.Add(oOcc)

                If Not oOcc.Suppressed AndAlso oOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Dim subOccs As ComponentOccurrences = Nothing
                    Try : subOccs = oOcc.SubOccurrences : Catch : subOccs = Nothing : End Try

                    If subOccs IsNot Nothing AndAlso subOccs.Count > 0 Then
                        ObtenerOcurrenciasRecursivas(subOccs, lista)
                    End If
                End If
            Catch
                Continue For
            End Try
        Next
    End Sub

    Private Function GenerarClaveUnica(ByVal oOcc As ComponentOccurrence) As String
        Dim sPath As String = oOcc.Name
        Dim parent As ComponentOccurrence = Nothing

        Try
            parent = oOcc.ParentOccurrence
        Catch
            parent = Nothing
        End Try

        While parent IsNot Nothing
            sPath = parent.Name & "\" & sPath
            Try
                parent = parent.ParentOccurrence
            Catch
                Exit While
            End Try
        End While

        Return sPath
    End Function

    Private Function TryGetDocumentFromOcc(ByVal oOcc As ComponentOccurrence, ByRef oDoc As Document) As Boolean
        Try
            Dim oTargetDef As ComponentDefinition = oOcc.Definition

            If TypeOf oTargetDef Is PartComponentDefinition Then
                Dim oPartDef As PartComponentDefinition = CType(oTargetDef, PartComponentDefinition)
                If oPartDef.IsModelStateMember Then
                    oDoc = CType(oPartDef.FactoryDocument, Document)
                Else
                    oDoc = oPartDef.Document
                End If
                Return True
            ElseIf TypeOf oTargetDef Is AssemblyComponentDefinition Then
                Dim oAsmDef As AssemblyComponentDefinition = CType(oTargetDef, AssemblyComponentDefinition)
                If oAsmDef.IsModelStateMember Then
                    oDoc = CType(oAsmDef.FactoryDocument, Document)
                Else
                    oDoc = oAsmDef.Document
                End If
                Return True
            End If
            Return False
        Catch
            Return False
        End Try
    End Function

    Private Function IsSpecialComponent(ByVal oOcc As ComponentOccurrence) As Boolean
        Try
            If oOcc.Suppressed Then Return True

            Dim oDef As ComponentDefinition = oOcc.Definition
            If oDef.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then Return True

            Dim oDoc As Document = Nothing
            If TryGetDocumentFromOcc(oOcc, oDoc) AndAlso oDoc IsNot Nothing Then
                Dim path As String = oDoc.FullFileName.ToLower()
                If path.Contains("content center") OrElse path.Contains("cc") OrElse path.Contains("libraries") Then
                    Return True
                End If
                ' Evitar factory documents (model states) como ocurrencias directas
                If TypeOf oDoc Is PartDocument Then
                    Dim oPartDef As PartComponentDefinition = CType(oDoc, PartDocument).ComponentDefinition
                    If oPartDef.IsModelStateMember Then Return True
                End If
            End If
            Return False
        Catch
            Return True
        End Try
    End Function

    Private Function EsRoutedSystem(ByVal oDoc As Document) As Boolean
        Try
            If oDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then Return False

            Dim oAsmDoc As AssemblyDocument = CType(oDoc, AssemblyDocument)
            Dim sName As String = IOPath.GetFileNameWithoutExtension(oDoc.FullFileName).ToLower()

            If sName.Contains("conducto") OrElse sName.Contains("conduit") OrElse sName.Contains("route") Then
                Return True
            End If

            Try
                Dim pipeDef As Object = oAsmDoc.ComponentDefinition
                If pipeDef.GetType().GetProperty("Routes") IsNot Nothing Then
                    Dim routes As Object = pipeDef.GetType().GetProperty("Routes").GetValue(pipeDef, Nothing)
                    If routes IsNot Nothing AndAlso CInt(routes.GetType().GetProperty("Count").GetValue(routes, Nothing)) > 0 Then
                        Return True
                    End If
                End If
            Catch
            End Try

        Catch
        End Try
        Return False
    End Function

    Private Function EsComponenteComercial(ByVal oDoc As Document) As Boolean
        Dim sComValue As String = GetPropertyValue(oDoc, "COM")
        If String.IsNullOrEmpty(sComValue) Then
            sComValue = GetCustomPropertyValue(oDoc, "COM")
        End If

        If String.IsNullOrEmpty(sComValue) Then Return False

        Dim sNormalized As String = sComValue.Trim().ToUpperInvariant()
        Return sNormalized = "✓" OrElse sNormalized = "TRUE" OrElse sNormalized = "YES" OrElse sNormalized = "1" OrElse sNormalized = "SI"
    End Function

    Private Function BuildNewName(ByVal oDoc As Document, ByVal sPrefix As String) As String
        Dim bEsCom As Boolean = EsComponenteComercial(oDoc)

        If bEsCom Then
            Dim sPartNumber As String = GetPropertyValue(oDoc, "Part Number")
            If String.IsNullOrEmpty(sPartNumber) Then
                sPartNumber = GetCustomPropertyValue(oDoc, "Part Number")
            End If

            If String.IsNullOrEmpty(sPartNumber) Then
                sPartNumber = GetPropertyValue(oDoc, "Stock Number")
                If String.IsNullOrEmpty(sPartNumber) Then
                    sPartNumber = GetCustomPropertyValue(oDoc, "Stock Number")
                End If
                If String.IsNullOrEmpty(sPartNumber) Then
                    sPartNumber = sPrefix
                End If
            End If

            RegistrarLog("   📋 COM detectado: " & IOPath.GetFileName(oDoc.FullFileName) & " -> " & sPartNumber)
            Return CleanFileName(sPartNumber, True)
        End If

        Dim sStockNumber As String = GetPropertyValue(oDoc, "Stock Number")
        Dim sPartNumberStd As String = GetPropertyValue(oDoc, "Part Number")

        If String.IsNullOrEmpty(sStockNumber) Then
            sStockNumber = GetCustomPropertyValue(oDoc, "Stock Number")
        End If
        If String.IsNullOrEmpty(sPartNumberStd) Then
            sPartNumberStd = GetCustomPropertyValue(oDoc, "Part Number")
        End If

        If String.IsNullOrEmpty(sStockNumber) Then
            sStockNumber = sPrefix
        End If

        If String.IsNullOrEmpty(sStockNumber) AndAlso String.IsNullOrEmpty(sPartNumberStd) Then
            Return ""
        End If

        Dim sFinalName As String = ""

        ' If Not String.IsNullOrEmpty(sStockNumber) AndAlso Not String.IsNullOrEmpty(sPartNumberStd) Then
        'sFinalName = sStockNumber & " - " & sPartNumberStd
        sFinalName = sPartNumberStd
        ' ElseIf Not String.IsNullOrEmpty(sStockNumber) Then
        '    sFinalName = sStockNumber
        'Else
        '   sFinalName = sPartNumberStd
        ' End If

        Return CleanFileName(sFinalName, False)
    End Function

    Private Function BuildMainAssemblyName(ByVal oDoc As Document, ByVal sPrefix As String) As String
        Dim sStockNumber As String = GetPropertyValue(oDoc, "Stock Number")
        Dim sPartNumber As String = GetPropertyValue(oDoc, "Part Number")
        Dim sProject As String = GetPropertyValue(oDoc, "Project")
        Dim sVendor As String = GetPropertyValue(oDoc, "Vendor")

        If String.IsNullOrEmpty(sStockNumber) Then
            sStockNumber = GetCustomPropertyValue(oDoc, "Stock Number")
        End If
        If String.IsNullOrEmpty(sPartNumber) Then
            sPartNumber = GetCustomPropertyValue(oDoc, "Part Number")
        End If
        If String.IsNullOrEmpty(sProject) Then
            sProject = GetCustomPropertyValue(oDoc, "Project")
        End If
        If String.IsNullOrEmpty(sVendor) Then
            sVendor = GetCustomPropertyValue(oDoc, "Vendor")
        End If

        If String.IsNullOrEmpty(sStockNumber) Then
            sStockNumber = sPrefix
        End If

        If String.IsNullOrEmpty(sStockNumber) AndAlso String.IsNullOrEmpty(sPartNumber) _
           AndAlso String.IsNullOrEmpty(sProject) AndAlso String.IsNullOrEmpty(sVendor) Then
            Return ""
        End If

        Dim parts As New List(Of String)

        If Not String.IsNullOrEmpty(sStockNumber) Then
            parts.Add(sStockNumber)
        End If
        If Not String.IsNullOrEmpty(sPartNumber) Then
            parts.Add(sPartNumber)
        End If
        If Not String.IsNullOrEmpty(sProject) Then
            parts.Add(sProject)
        End If
        If Not String.IsNullOrEmpty(sVendor) Then
            parts.Add(sVendor)
        End If

        'Dim sFinalName As String = String.Join(" - ", parts)

        Dim sFinalName As String = sPartNumber
        Return CleanFileName(sFinalName, False)
    End Function

    Private Function GetPropertyValue(ByVal oDoc As Document, ByVal sPropName As String) As String
        Dim val As String = ""
        Try : val = oDoc.PropertySets.Item("Design Tracking Properties").Item(sPropName).Value.ToString().Trim() : Catch : End Try
        If String.IsNullOrEmpty(val) Then
            Try : val = oDoc.PropertySets.Item("Inventor Summary Information").Item(sPropName).Value.ToString().Trim() : Catch : End Try
        End If
        Return val
    End Function

    Private Function GetCustomPropertyValue(ByVal oDoc As Document, ByVal sPropName As String) As String
        Dim val As String = ""
        Try : val = oDoc.PropertySets.Item("Inventor User Defined Properties").Item(sPropName).Value.ToString().Trim() : Catch : End Try
        Return val
    End Function

    Private Function CleanFileName(ByVal sInput As String, ByVal bAggressive As Boolean) As String
        If String.IsNullOrEmpty(sInput) Then Return ""

        Dim sResult As String = sInput

        sResult = sResult.Replace(vbCr, " ").Replace(vbLf, " ")
        sResult = Regex.Replace(sResult, "\s+", " ")

        Dim sInvalid As String = "\/:?*""<>|"
        For Each c As Char In sInvalid
            sResult = sResult.Replace(c, "-"c)
        Next

        sResult = sResult.Replace("Ø", "DIA").Replace("Æ", "AE").Replace("Ñ", "N")

        If bAggressive Then
            Dim aggressiveChars() As Char = {","c, "."c, ";"c, "="c, "+"c, "@"c, "#"c, "$"c, "%"c, "&"c, "'"c, "("c, ")"c, "["c, "]"c, "{"c, "}"c, "<"c, ">"c}
            For Each c As Char In aggressiveChars
                sResult = sResult.Replace(c, "-"c)
            Next

            sResult = Regex.Replace(sResult, "-{2,}", "-")
            sResult = sResult.Trim("-"c)
            If sResult.Length > 60 Then
                sResult = sResult.Substring(0, 60).Trim("-"c)
            End If
        Else
            sResult = Regex.Replace(sResult, "-{2,}", "-")
            sResult = sResult.Trim("-"c)
            If sResult.Length > 100 Then
                sResult = sResult.Substring(0, 100).Trim("-"c)
            End If
        End If

        Return sResult.Trim()
    End Function

    Private Sub ProcessMainAssembly(ByVal oAsmDoc As AssemblyDocument, ByVal sPrefix As String, ByVal sAsmOriginalPath As String)
        Try
            Dim sNewName As String = BuildMainAssemblyName(oAsmDoc, sPrefix)
            If String.IsNullOrEmpty(sNewName) Then Return

            Dim sFolder As String = IOPath.GetDirectoryName(sAsmOriginalPath)
            Dim sExt As String = IOPath.GetExtension(sAsmOriginalPath)
            Dim sNewFullPath As String = IOPath.Combine(sFolder, sNewName & sExt)

            If String.Equals(sAsmOriginalPath, sNewFullPath, StringComparison.OrdinalIgnoreCase) Then Return

            RegistrarLog(vbCrLf & "👑 RAÍZ: " & IOPath.GetFileName(sAsmOriginalPath) & " -> " & sNewName & sExt)

            If Not bDryRun Then
                ' FIX: Forzar update antes de SaveAs para evitar inconsistencias
                Try : oAsmDoc.Update2(True) : Catch : End Try
                oAsmDoc.SaveAs(sNewFullPath, False)
                If Not String.Equals(sAsmOriginalPath, sNewFullPath, StringComparison.OrdinalIgnoreCase) Then
                    _filesToDelete.Add(sAsmOriginalPath)
                End If
            End If
            iProcessed += 1
        Catch ex As Exception
            RegistrarLog("❌ ERROR raíz: " & ex.Message)
            iErrors += 1
        End Try
    End Sub

    Private Function EliminarArchivosOriginales() As Integer
        If _hasRoutedSystems Then
            RegistrarLog("⚠️ NO SE ELIMINARON ORIGINALES: Se detectaron Routed Systems en el modelo.")
            RegistrarLog("   Elimine manualmente los archivos originales si es necesario, o ejecute sin routed systems." & vbCrLf)
            Return 0
        End If

        Dim iDeleted As Integer = 0
        Dim iFailed As Integer = 0

        For Each sOriginalPath As String In _filesToDelete
            If Not _renamedFiles.ContainsKey(sOriginalPath) Then Continue For

            Dim sClonedPath As String = _renamedFiles(sOriginalPath)
            If Not IOFile.Exists(sClonedPath) Then
                RegistrarLog("⚠️ No se elimina original (clon no existe): " & IOPath.GetFileName(sOriginalPath))
                iFailed += 1
                Continue For
            End If

            Try
                For Each oDoc As Document In oApp.Documents
                    If String.Equals(oDoc.FullFileName, sOriginalPath, StringComparison.OrdinalIgnoreCase) Then
                        Try
                            oDoc.Close(False)
                            RegistrarLog("   📁 Cerrado en Inventor: " & IOPath.GetFileName(sOriginalPath))
                        Catch exClose As Exception
                            RegistrarLog("   ⚠️ No se pudo cerrar: " & IOPath.GetFileName(sOriginalPath) & " - " & exClose.Message)
                        End Try
                        Exit For
                    End If
                Next
            Catch
            End Try

            Dim bDeleted As Boolean = False
            Dim iAttempts As Integer = 0
            While iAttempts < 3 AndAlso Not bDeleted
                iAttempts += 1
                Try
                    If iAttempts > 1 Then
                        System.Threading.Thread.Sleep(500)
                    End If

                    IOFile.Delete(sOriginalPath)
                    bDeleted = True
                    RegistrarLog("🗑️ Eliminado: " & IOPath.GetFileName(sOriginalPath))
                    iDeleted += 1

                Catch ex As Exception
                    If iAttempts >= 3 Then
                        RegistrarLog("❌ No se pudo eliminar: " & IOPath.GetFileName(sOriginalPath) & " - " & ex.Message)
                        iFailed += 1
                    End If
                End Try
            End While
        Next

        RegistrarLog(vbCrLf & "Resumen eliminación: " & iDeleted & " eliminados, " & iFailed & " fallidos")
        Return iDeleted
    End Function

    Private Sub MostrarDialogoFinalizacion(ByVal iDeleted As Integer)
        Try
            Dim sModo As String = If(bDryRun, "MODO SIMULACIÓN (DryRun)", "MODO EJECUCIÓN")

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("═══════════════════════════════════════")
            sb.AppendLine("   RENOMBRADOR - PROCESO FINALIZADO")
            sb.AppendLine("═══════════════════════════════════════")
            sb.AppendLine()
            sb.AppendLine(sModo)
            sb.AppendLine()
            sb.AppendLine("📊 RESUMEN:")
            sb.AppendLine("   • Procesados:  " & iProcessed.ToString())
            sb.AppendLine("   • Omitidos:    " & iSkipped.ToString())
            sb.AppendLine("   • Errores:     " & iErrors.ToString())
            If Not bDryRun Then
                sb.AppendLine("   • Eliminados:  " & iDeleted.ToString())
            End If
            If _hasRoutedSystems Then
                sb.AppendLine()
                sb.AppendLine("⚠️ ATENCIÓN: Se detectaron Routed Systems.")
                sb.AppendLine("   Los archivos originales NO fueron eliminados")
                sb.AppendLine("   para evitar referencias rotas.")
                sb.AppendLine("   Todos los componentes de routed systems fueron")
                sb.AppendLine("   completamente ignorados (no clonados ni reemplazados).")
            End If
            sb.AppendLine()
            sb.AppendLine("📁 Log guardado en:")
            sb.AppendLine("   " & sLogFilePath)

            If iErrors > 0 Then
                sb.AppendLine()
                sb.AppendLine("⚠️ Se detectaron errores durante el proceso.")
                sb.AppendLine("   Revisá el log para más detalles.")
            End If

            Dim icon As MessageBoxIcon
            Dim caption As String

            If iErrors > 0 Then
                icon = MessageBoxIcon.Warning
                caption = "Proceso finalizado con advertencias"
            ElseIf bDryRun Then
                icon = MessageBoxIcon.Information
                caption = "Simulación completada"
            Else
                icon = MessageBoxIcon.Information
                caption = "Proceso completado exitosamente"
            End If

            MessageBox.Show(sb.ToString(), caption, MessageBoxButtons.OK, icon)

        Catch ex As Exception
            RegistrarLog("❌ Error mostrando diálogo: " & ex.Message)
        End Try
    End Sub

    Private Sub RegistrarLog(ByVal mensaje As String)
        Dim timestamp As String = DateTime.Now.ToString("HH:mm:ss.fff")
        sLog &= timestamp & " | " & mensaje & vbCrLf
        RaiseEvent ProgresoReportado(mensaje)
    End Sub

End Class