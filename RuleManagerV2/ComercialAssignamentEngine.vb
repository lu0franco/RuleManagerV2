Imports Inventor
Imports System.Linq
Imports System.Collections.Generic

''' <summary>
''' Motor robusto de asignación de propiedad COM.
''' Reemplaza la regla iLogic "Asignar Comerciales.iLogicVb"
''' </summary>
Public Class ComercialAssignmentEngine

    Private ReadOnly _app As Inventor.Application
    Private ReadOnly _log As List(Of String)
    Private ReadOnly _processedFiles As HashSet(Of String)

    Public Sub New(app As Inventor.Application)
        _app = app
        _log = New List(Of String)
        _processedFiles = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    End Sub

    ''' <summary>
    ''' Ejecuta la asignación completa desde el documento activo (modo automático).
    ''' </summary>
    Public Function ExecuteFromActiveDocument() As AssignmentResult
        Dim doc As Document = _app.ActiveDocument
        If doc Is Nothing Then
            Return AssignmentResult.Failed("No hay documento activo")
        End If

        If doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
            Return AssignmentResult.Failed("El documento activo no es un ensamblaje (.iam)")
        End If

        Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
        Return Execute(asmDoc)
    End Function

    ''' <summary>
    ''' Ejecuta la asignación sobre un ensamblaje específico (modo automático).
    ''' </summary>
    Public Function Execute(asmDoc As AssemblyDocument) As AssignmentResult
        _log.Clear()
        _processedFiles.Clear()

        Log("=== INICIO ASIGNACIÓN COM (automático) ===")
        Log("Ensamblaje: " & asmDoc.FullFileName)

        Dim resultado As New AssignmentResult()

        Try
            Dim todasLasOcurrencias = ObtenerTodasLasOcurrencias(asmDoc)
            Log("Total ocurrencias: " & todasLasOcurrencias.Count)

            Dim comerciales As New List(Of ComponentOccurrence)
            Dim noComerciales As New List(Of ComponentOccurrence)

            For Each occ In todasLasOcurrencias
                If EsContentCenter(occ) OrElse EsBuloneria(occ) Then
                    comerciales.Add(occ)
                Else
                    noComerciales.Add(occ)
                End If
            Next

            Log("A limpiar (CC/Bulonería): " & comerciales.Count)
            Log("A marcar COM: " & noComerciales.Count)

            Dim okCount As Integer = 0
            Dim failCount As Integer = 0

            For Each occ In noComerciales
                If ProcesarOcurrencia(occ, "✓", resultado) Then
                    okCount += 1
                Else
                    failCount += 1
                End If
            Next

            For Each occ In comerciales
                If ProcesarOcurrencia(occ, "", resultado) Then
                    okCount += 1
                Else
                    failCount += 1
                End If
            Next

            Try
                asmDoc.Update2(True)
                _app.ActiveView.Update()
            Catch ex As Exception
                Log("WARN: Error al refrescar vista: " & ex.Message)
            End Try

            resultado.Success = True
            resultado.ProcessedCount = okCount
            resultado.FailedCount = failCount
            resultado.LogEntries = _log.ToList()

            Log("=== FIN OK | Procesados: " & okCount & " | Fallidos: " & failCount & " ===")

        Catch ex As Exception
            resultado.Success = False
            resultado.ErrorMessage = ex.Message
            Log("=== ERROR CRÍTICO: " & ex.Message & " ===")
        End Try

        Return resultado
    End Function

    ''' <summary>
    ''' Ejecuta la asignación usando cambios definidos manualmente (desde el diálogo).
    ''' </summary>
    Public Function ExecuteWithChanges(cambios As List(Of NodoComercial)) As AssignmentResult
        _log.Clear()
        _processedFiles.Clear()

        Log("=== INICIO ASIGNACIÓN COM (desde diálogo) ===")

        Dim resultado As New AssignmentResult()
        Dim okCount As Integer = 0
        Dim failCount As Integer = 0

        Try
            For Each nodo In cambios.Where(Function(n) n.TieneCambio AndAlso Not n.IsBroken)
                Dim success = EscribirPropiedadCOM(nodo.FileName, nodo.ComNuevo)
                If success Then
                    okCount += 1
                    Log("OK [" & If(nodo.ComNuevo = "", "LIMPIO", nodo.ComNuevo) & "]: " & System.IO.Path.GetFileName(nodo.FileName))
                Else
                    failCount += 1
                    Log("FAIL: " & System.IO.Path.GetFileName(nodo.FileName))
                End If
            Next

            Try
                Dim doc As Document = _app.ActiveDocument
                If doc IsNot Nothing Then
                    doc.Update2(True)
                    _app.ActiveView.Update()
                End If
            Catch ex As Exception
                Log("WARN: Error al refrescar: " & ex.Message)
            End Try

            resultado.Success = True
            resultado.ProcessedCount = okCount
            resultado.FailedCount = failCount
            resultado.LogEntries = _log.ToList()

            Log("=== FIN | Procesados: " & okCount & " | Fallidos: " & failCount & " ===")

        Catch ex As Exception
            resultado.Success = False
            resultado.ErrorMessage = ex.Message
            Log("=== ERROR: " & ex.Message & " ===")
        End Try

        Return resultado
    End Function

    ' =========================================================
    ' MÉTODOS PRIVADOS
    ' =========================================================

    Private Function ObtenerTodasLasOcurrencias(asmDoc As AssemblyDocument) As List(Of ComponentOccurrence)
        Dim lista As New List(Of ComponentOccurrence)
        RecorrerOcurrencias(asmDoc.ComponentDefinition.Occurrences, lista)
        Return lista
    End Function

    Private Sub RecorrerOcurrencias(occs As ComponentOccurrences, lista As List(Of ComponentOccurrence))
        For Each occ As ComponentOccurrence In occs
            lista.Add(occ)
            If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                RecorrerOcurrencias(occ.SubOccurrences, lista)
            End If
        Next
    End Sub

    Private Function EsContentCenter(occ As ComponentOccurrence) As Boolean
        Try
            If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                Return occ.Definition.IsContentMember
            End If
        Catch
        End Try
        Return False
    End Function

    Private Function EsBuloneria(occ As ComponentOccurrence) As Boolean
        Try
            If occ.Definition Is Nothing OrElse occ.Definition.Document Is Nothing Then Return False
            Dim doc As Document = occ.Definition.Document
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

    Private Function ProcesarOcurrencia(occ As ComponentOccurrence, valor As String, resultado As AssignmentResult) As Boolean
        Try
            If occ Is Nothing Then
                Log("SKIP: Ocurrencia es Nothing")
                Return False
            End If

            If occ.Definition Is Nothing Then
                Log("SKIP: " & occ.Name & " | Definition es Nothing")
                resultado.BrokenReferences.Add(occ.Name)
                Return False
            End If

            If occ.Definition.Document Is Nothing Then
                Log("SKIP: " & occ.Name & " | Definition.Document es Nothing")
                Return False
            End If

            Dim fileName As String = occ.Definition.Document.FullDocumentName
            If String.IsNullOrWhiteSpace(fileName) Then
                Log("SKIP: " & occ.Name & " | FullDocumentName vacío")
                Return False
            End If

            If _processedFiles.Contains(fileName) Then
                Return True
            End If

            Dim success = EscribirPropiedadCOM(fileName, valor)
            If success Then
                _processedFiles.Add(fileName)
                Log("OK [" & If(valor = "", "LIMPIO", valor) & "]: " & System.IO.Path.GetFileName(fileName))
                Return True
            Else
                Log("FAIL: " & System.IO.Path.GetFileName(fileName))
                Return False
            End If

        Catch ex As Exception
            Log("EXCEPCIÓN en " & If(occ?.Name, "?") & ": " & ex.Message)
            Return False
        End Try
    End Function

    Private Function EscribirPropiedadCOM(fileName As String, valor As String) As Boolean
        Dim doc As Document = Nothing
        Dim fueAbiertoPorMi As Boolean = False

        Try
            doc = BuscarDocumentoAbierto(fileName)

            If doc Is Nothing Then
                doc = _app.Documents.Open(fileName, False)
                fueAbiertoPorMi = True
            End If

            If doc Is Nothing Then
                Log("  -> No se pudo obtener documento")
                Return False
            End If

            Dim userProps As PropertySet
            Try
                userProps = doc.PropertySets.Item("Inventor User Defined Properties")
            Catch
                userProps = doc.PropertySets.Add("Inventor User Defined Properties")
            End Try

            Dim propCOM As Inventor.Property = Nothing
            For Each p As Inventor.Property In userProps
                If p.Name = "COM" Then
                    propCOM = p
                    Exit For
                End If
            Next

            If propCOM Is Nothing Then
                userProps.Add(valor, "COM")
            Else
                propCOM.Value = valor
            End If

            doc.Save2(False)

            If fueAbiertoPorMi Then
                doc.Close(True)
            End If

            Return True

        Catch ex As Exception
            Log("  -> ERROR escribiendo COM: " & ex.Message)
            If doc IsNot Nothing AndAlso fueAbiertoPorMi Then
                Try
                    doc.Close(False)
                Catch
                End Try
            End If
            Return False
        End Try
    End Function

    Private Function BuscarDocumentoAbierto(fileName As String) As Document
        Try
            For Each doc As Document In _app.Documents
                If doc.FullFileName.Equals(fileName, StringComparison.OrdinalIgnoreCase) Then
                    Return doc
                End If
            Next
        Catch
        End Try
        Return Nothing
    End Function

    Private Sub Log(msg As String)
        Dim timestamp As String = DateTime.Now.ToString("HH:mm:ss.fff")
        _log.Add("[" & timestamp & "] " & msg)
    End Sub

End Class

''' <summary>
''' Resultado de la operación de asignación.
''' </summary>
Public Class AssignmentResult
    Public Property Success As Boolean
    Public Property ProcessedCount As Integer
    Public Property FailedCount As Integer
    Public Property ErrorMessage As String
    Public Property LogEntries As List(Of String)
    Public Property BrokenReferences As New List(Of String)

    Public Shared Function Failed(msg As String) As AssignmentResult
        Return New AssignmentResult With {
            .Success = False,
            .ErrorMessage = msg,
            .LogEntries = New List(Of String) From {msg}
        }
    End Function

    Public Sub New()
        LogEntries = New List(Of String)
        BrokenReferences = New List(Of String)
    End Sub
End Class