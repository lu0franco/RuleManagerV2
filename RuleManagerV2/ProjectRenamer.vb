Imports Inventor
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Linq

Imports IOPath = System.IO.Path
Imports IOFile = System.IO.File
Imports IODirectory = System.IO.Directory

Public Class ProjectRenamer

    Private ReadOnly _app As Inventor.Application
    Private ReadOnly _asmDoc As AssemblyDocument
    Private _logPath As String

    Public Sub New(app As Inventor.Application, asmDoc As AssemblyDocument)
        _app = app
        _asmDoc = asmDoc
    End Sub

    Public Function Execute() As String

        _logPath = "C:\Temp\ProjectRenamer.log"

        If IOFile.Exists(_logPath) Then
            IOFile.Delete(_logPath)
        End If

        _app.SilentOperation = True
        _app.UserInterfaceManager.UserInteractionDisabled = True

        Try
            Log("===== START =====")

            ' =========================================================
            ' PASO 0: GUARDAR DOCUMENTO ORIGINAL CON CAMBIOS RECIENTES
            ' =========================================================
            Try
                _asmDoc.Save2(False)
                Log("Documento original guardado")
            Catch ex As Exception
                Log("ADVERTENCIA: No se pudo guardar documento original: " & ex.Message)
            End Try

            ' =========================================================
            ' PASO 1: EXTRAER DATOS DEL DOCUMENTO ACTIVO (DATOS FRESCOS)
            ' =========================================================
            Dim data As ProjectData = ExtractProjectDataFromDoc(_asmDoc)
            Dim newProjectName As String = BuildProjectName(data)

            Log("Nuevo proyecto (inicial): " & newProjectName)

            ' Guardamos las rutas del proyecto de origen
            Dim currentProject As DesignProject = _app.DesignProjectManager.ActiveDesignProject
            Dim sourceFolder As String = currentProject.WorkspacePath
            Dim parentFolder As String = Directory.GetParent(sourceFolder).FullName
            Dim targetFolder As String = IOPath.Combine(parentFolder, newProjectName)

            If IODirectory.Exists(targetFolder) Then
                Throw New Exception("La carpeta destino ya existe.")
            End If

            ' CAPTURA DINAMICA: Nombre exacto del IAM abierto actualmente
            Dim oldAsmName As String = IOPath.GetFileName(_asmDoc.FullFileName)
            Dim oldAsmNameNoExt As String = IOPath.GetFileNameWithoutExtension(oldAsmName)
            Log("IAM original a reemplazar detectado como: " & oldAsmName)

            CloseAllDocuments()

            ' =========================================================
            ' PASO 1b: LIBERAR PROYECTO ACTIVO PARA EVITAR LOCK DEL IPJ
            ' =========================================================
            Try
                Dim mgr As DesignProjectManager = _app.DesignProjectManager
                Dim defaultProj As DesignProject = Nothing

                ' Buscar un proyecto default seguro
                For Each proj As DesignProject In mgr.DesignProjects
                    If proj.Name.ToLower().Contains("default") Then
                        defaultProj = proj
                        Exit For
                    End If
                Next

                ' Fallback: usar cualquier proyecto que no sea el actual
                If defaultProj Is Nothing AndAlso mgr.DesignProjects.Count > 0 Then
                    For Each proj As DesignProject In mgr.DesignProjects
                        If Not String.Equals(proj.FullFileName, currentProject.FullFileName, StringComparison.OrdinalIgnoreCase) Then
                            defaultProj = proj
                            Exit For
                        End If
                    Next
                End If

                If defaultProj IsNot Nothing Then
                    defaultProj.Activate()
                    Log("Proyecto temporal activado: " & defaultProj.Name)
                    System.Threading.Thread.Sleep(1000)
                Else
                    Log("ADVERTENCIA: No se encontro proyecto temporal para activar")
                End If

                ' Liberar referencia COM del proyecto anterior
                Marshal.ReleaseComObject(currentProject)
                System.Threading.Thread.Sleep(500)

            Catch ex As Exception
                Log("ADVERTENCIA al liberar proyecto: " & ex.Message)
            End Try

            '====================================================
            ' 2. COPIAR TODO EL PROYECTO A DESTINO
            '====================================================
            Log("Copiando carpeta base completa")
            CopyDirectory(sourceFolder, targetFolder)

            ' =========================================================
            ' PASO 2b: FORZAR LIBERACION DE HANDLES ANTES DE TOCAR IPJ
            ' =========================================================
            For gcCycle As Integer = 1 To 3
                GC.Collect()
                GC.WaitForPendingFinalizers()
                GC.Collect()
                System.Threading.Thread.Sleep(500)
            Next

            '====================================================
            ' 3. RENOMBRAR IPJ EN DESTINO (antes de activarlo)
            '====================================================
            Dim oldIpjPath As String = IODirectory.GetFiles(targetFolder, "*.ipj")(0)
            Dim newIpjPath As String = IOPath.Combine(targetFolder, newProjectName & ".ipj")

            Log("Renombrando IPJ en destino: " & IOPath.GetFileName(oldIpjPath) & " -> " & IOPath.GetFileName(newIpjPath))
            IOFile.Move(oldIpjPath, newIpjPath)

            '====================================================
            ' 4. CREAR EL NUEVO IAM MASTER (Manteniendo el viejo en disco)
            '====================================================
            Dim copiedOldAsmPath As String = FindFile(targetFolder, oldAsmName)
            If String.IsNullOrEmpty(copiedOldAsmPath) Then
                Throw New Exception("No se encontro el IAM original (" & oldAsmName & ") en la carpeta copiada.")
            End If

            Dim newAsmPath As String = IOPath.Combine(IOPath.GetDirectoryName(copiedOldAsmPath), newProjectName & ".iam")

            Log("Abriendo IAM Master copiado")
            Dim asmCopy As AssemblyDocument = CType(_app.Documents.Open(copiedOldAsmPath, False), AssemblyDocument)

            ' =========================================================
            ' 4b. COPIAR PROPIEDADES iPROPERTY DEL ORIGINAL AL NUEVO
            ' =========================================================
            Log("Copiando propiedades iProperty del original al nuevo IAM...")
            CopyPropertiesBetweenDocuments(_asmDoc, asmCopy)

            Log("Generando copia de IAM Master con nombre nuevo")
            asmCopy.SaveAs(newAsmPath, True)
            asmCopy.Close(True)
            Marshal.ReleaseComObject(asmCopy)

            ' =========================================================
            ' 5. ACTIVAR PROYECTO NUEVO (recien ahora, despues de renombrar todo)
            ' =========================================================
            Log("Activando proyecto nuevo")
            ActivateProject(newIpjPath)

            '====================================================
            ' 6. ACTUALIZAR REFERENCIAS EN SUBENSAMBLAJES E IPTs
            '====================================================
            Log("Actualizando referencias en documentos dependientes...")
            UpdateReferencesInAllDocuments(targetFolder, oldAsmName, newAsmPath)

            '====================================================
            ' 7. BUSCAR Y PROCESAR IDWs EN LA CARPETA NUEVA
            '====================================================
            Log("Buscando IDWs copiados en carpeta destino")
            Dim allIdws = IODirectory.GetFiles(targetFolder, "*.idw", SearchOption.AllDirectories) _
                    .Where(Function(f) Not f.ToLower().Contains("\oldversions")) _
                    .Select(Function(f) New FileInfo(f)) _
                    .OrderByDescending(Function(fi) fi.Length) _
                    .ToList()

            For Each idwInfo As FileInfo In allIdws
                Dim oldIdwPath As String = idwInfo.FullName
                Dim oldBase As String = IOPath.GetFileNameWithoutExtension(oldIdwPath)

                Log("Procesando IDW: " & oldBase)

                Try
                    ' Abrimos el plano en segundo plano
                    Dim doc As DrawingDocument = CType(_app.Documents.Open(oldIdwPath, False), DrawingDocument)

                    Dim fds As Object = doc.File.ReferencedFileDescriptors
                    Log("-> Referencias encontradas en este IDW: " & fds.Count)

                    Dim replaced As Boolean = False

                    ' Bucle indexado numerico base 1
                    For i As Integer = 1 To fds.Count
                        Dim fd As FileDescriptor = CType(fds.Item(i), FileDescriptor)
                        Dim currentRefName As String = ""
                        Dim currentFullPath As String = ""

                        Try
                            currentRefName = IOPath.GetFileName(fd.LogicalFileName)
                        Catch
                        End Try

                        If String.IsNullOrEmpty(currentRefName) Then
                            Try
                                currentRefName = IOPath.GetFileName(fd.FullFileName)
                            Catch
                            End Try
                        End If

                        Try
                            currentFullPath = fd.FullFileName
                        Catch
                        End Try

                        ' Solo procesar referencias .iam
                        Dim currentRefExt As String = IOPath.GetExtension(currentRefName).ToLower()
                        If currentRefExt <> ".iam" Then
                            Continue For
                        End If

                        ' Solo reemplazar si es EXACTAMENTE el IAM master viejo
                        Dim isExactMatch As Boolean = String.Equals(currentRefName, oldAsmName, StringComparison.OrdinalIgnoreCase)
                        Dim isFullPathMatch As Boolean = Not String.IsNullOrEmpty(currentFullPath) AndAlso
                            String.Equals(IOPath.GetFileName(currentFullPath), oldAsmName, StringComparison.OrdinalIgnoreCase)

                        If isExactMatch OrElse isFullPathMatch Then
                            Log("   -> [COINCIDENCIA] Reemplazando referencia a: " & currentRefName)
                            If Not String.IsNullOrEmpty(currentFullPath) Then
                                Log("      Full path: " & currentFullPath)
                            End If
                            Log("      Por: " & newAsmPath)

                            fd.ReplaceReference(newAsmPath)
                            replaced = True
                        Else
                            If Not String.IsNullOrEmpty(currentRefName) Then
                                Log("   -> [DEBUG] Ref: " & currentRefName & " | Full: " & currentFullPath)
                            End If
                        End If
                    Next

                    If replaced Then
                        Log("-> Forzando actualizacion de la geometria (Update2)")
                        doc.Update2(True)
                        Log("-> Guardando cambios con Save2")
                        doc.Save2(False)
                    Else
                        Log("-> [ADVERTENCIA] No se reemplazo nada. Ninguna referencia interna coincidio con: " & oldAsmName)
                    End If

                    Log("-> Cerrando IDW")
                    doc.Close(True)
                    Marshal.ReleaseComObject(doc)
                    doc = Nothing

                Catch ex As Exception
                    Log("ERROR procesando IDW " & oldBase & ": " & ex.Message)
                End Try
            Next

            '====================================================
            ' 8. LIMPIEZA FINAL: RENOMBRAR IAM VIEJO A .OLD (NO BORRAR)
            '====================================================
            If IOFile.Exists(copiedOldAsmPath) Then
                Dim oldBackupPath As String = copiedOldAsmPath & ".old"
                Log("Limpieza: Renombrando IAM viejo a .old (no borrar para preservar referencias)")
                Try
                    If IOFile.Exists(oldBackupPath) Then
                        IOFile.Delete(oldBackupPath)
                    End If
                    IOFile.Move(copiedOldAsmPath, oldBackupPath)
                Catch ex As Exception
                    Log("ADVERTENCIA: No se pudo renombrar IAM viejo: " & ex.Message)
                End Try
            End If

            Log("===== END OK =====")
            Log("Proyecto activo final: " & newProjectName)

            Return targetFolder

        Catch ex As Exception
            Log("ERROR CRITICO: " & ex.Message)
            Throw
        Finally
            _app.UserInterfaceManager.UserInteractionDisabled = False
            _app.SilentOperation = False
        End Try

    End Function

    ' =========================================================
    ' COPIAR PROPIEDADES iPROPERTY ENTRE DOCUMENTOS
    ' =========================================================
    Private Sub CopyPropertiesBetweenDocuments(sourceDoc As Document, targetDoc As Document)
        Try
            ' Propiedades a copiar
            Dim propsToCopy As String() = {"Stock Number", "Part Number", "Project", "Vendor", "Description", "Revision Number"}

            For Each propName As String In propsToCopy
                Try
                    Dim sourceValue As String = ""
                    Try
                        sourceValue = sourceDoc.PropertySets.Item("Design Tracking Properties").Item(propName).Value.ToString()
                    Catch
                        ' Intentar en Summary Information
                        Try
                            sourceValue = sourceDoc.PropertySets.Item("Inventor Summary Information").Item(propName).Value.ToString()
                        Catch
                            Continue For
                        End Try
                    End Try

                    ' Escribir en destino
                    Try
                        targetDoc.PropertySets.Item("Design Tracking Properties").Item(propName).Value = sourceValue
                        Log("   -> Propiedad copiada: " & propName & " = " & sourceValue)
                    Catch
                        ' Si no existe, intentar crearla
                        Try
                            targetDoc.PropertySets.Item("Design Tracking Properties").Add(sourceValue, propName)
                            Log("   -> Propiedad creada: " & propName & " = " & sourceValue)
                        Catch exInner As Exception
                            Log("   -> No se pudo copiar propiedad " & propName & ": " & exInner.Message)
                        End Try
                    End Try

                Catch ex As Exception
                    Log("   -> Error copiando " & propName & ": " & ex.Message)
                End Try
            Next

            ' Guardar cambios en el documento destino
            targetDoc.Save2(False)
            Log("Propiedades copiadas y guardadas en destino")

        Catch ex As Exception
            Log("ADVERTENCIA al copiar propiedades: " & ex.Message)
        End Try
    End Sub

    ' =========================================================
    ' ACTUALIZAR REFERENCIAS EN TODOS LOS DOCUMENTOS DEL PROYECTO
    ' =========================================================
    Private Sub UpdateReferencesInAllDocuments(targetFolder As String, oldAsmName As String, newAsmPath As String)
        Dim allIamFiles = IODirectory.GetFiles(targetFolder, "*.iam", SearchOption.AllDirectories) _
            .Where(Function(f) Not f.ToLower().Contains("\oldversions")) _
            .ToList()

        Dim allIptFiles = IODirectory.GetFiles(targetFolder, "*.ipt", SearchOption.AllDirectories) _
            .Where(Function(f) Not f.ToLower().Contains("\oldversions")) _
            .ToList()

        ' Procesar IAMs
        For Each iamFile In allIamFiles
            ' No procesar el nuevo IAM master
            If String.Equals(IOPath.GetFileName(iamFile), IOPath.GetFileName(newAsmPath), StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            Try
                Dim doc As Document = _app.Documents.Open(iamFile, False)
                Dim fds As Object = doc.File.ReferencedFileDescriptors
                Dim replaced As Boolean = False

                For i As Integer = 1 To fds.Count
                    Dim fd As FileDescriptor = CType(fds.Item(i), FileDescriptor)
                    Dim refName As String = ""

                    Try
                        refName = IOPath.GetFileName(fd.FullFileName)
                    Catch
                        Try
                            refName = IOPath.GetFileName(fd.LogicalFileName)
                        Catch
                        End Try
                    End Try

                    If String.Equals(refName, oldAsmName, StringComparison.OrdinalIgnoreCase) Then
                        Log("   -> [SUB-REF] Reemplazando en " & IOPath.GetFileName(iamFile) & ": " & refName & " -> " & IOPath.GetFileName(newAsmPath))
                        fd.ReplaceReference(newAsmPath)
                        replaced = True
                    End If
                Next

                If replaced Then
                    doc.Save2(False)
                    Log("   -> [SUB-REF] Guardado: " & IOPath.GetFileName(iamFile))
                End If

                doc.Close(True)
                Marshal.ReleaseComObject(doc)

            Catch ex As Exception
                Log("   -> [SUB-REF ERROR] " & IOPath.GetFileName(iamFile) & ": " & ex.Message)
            End Try
        Next

        ' Procesar IPTs (por si alguna pieza referencia al IAM, aunque es raro)
        For Each iptFile In allIptFiles
            Try
                Dim doc As Document = _app.Documents.Open(iptFile, False)
                Dim fds As Object = doc.File.ReferencedFileDescriptors
                Dim replaced As Boolean = False

                For i As Integer = 1 To fds.Count
                    Dim fd As FileDescriptor = CType(fds.Item(i), FileDescriptor)
                    Dim refName As String = ""

                    Try
                        refName = IOPath.GetFileName(fd.FullFileName)
                    Catch
                        Try
                            refName = IOPath.GetFileName(fd.LogicalFileName)
                        Catch
                        End Try
                    End Try

                    If String.Equals(refName, oldAsmName, StringComparison.OrdinalIgnoreCase) Then
                        fd.ReplaceReference(newAsmPath)
                        replaced = True
                    End If
                Next

                If replaced Then
                    doc.Save2(False)
                End If

                doc.Close(True)
                Marshal.ReleaseComObject(doc)

            Catch
                ' Ignorar errores en IPTs
            End Try
        Next
    End Sub

    ' =========================================================
    ' EXTRAER DATOS DE UN DOCUMENTO
    ' =========================================================
    Private Function ExtractProjectDataFromDoc(doc As Document) As ProjectData
        Dim data As New ProjectData()
        data.Modelo = NormalizeCode(GetPropertyFromDoc(doc, "Stock Number"))
        data.Maquina = GetPropertyFromDoc(doc, "Part Number")
        data.OT = GetPropertyFromDoc(doc, "Project")
        data.Cliente = GetPropertyFromDoc(doc, "Vendor")
        Return data
    End Function

    Private Function GetPropertyFromDoc(doc As Document, propName As String) As String
        Try
            Return doc.PropertySets.Item("Design Tracking Properties").Item(propName).Value.ToString()
        Catch
            Return ""
        End Try
    End Function

    ' =========================================================
    ' METODOS AUXILIARES
    ' =========================================================
    Private Sub Log(msg As String)
        IOFile.AppendAllText(_logPath, DateTime.Now.ToString("HH:mm:ss") & " | " & msg & vbCrLf)
    End Sub

    Private Function FindFile(rootFolder As String, fileName As String) As String
        Dim files() As String = IODirectory.GetFiles(rootFolder, fileName, SearchOption.AllDirectories)
        For Each f As String In files
            If Not f.ToLower().Contains("\oldversions\") Then
                Return f
            End If
        Next
        Return Nothing
    End Function

    Private Sub CloseAllDocuments()
        Dim closedCount As Integer = 0
        Dim maxAttempts As Integer = 5

        For attempt As Integer = 1 To maxAttempts
            Dim docsToClose As New List(Of Document)

            Try
                For Each doc As Document In _app.Documents
                    If doc IsNot Nothing Then
                        docsToClose.Add(doc)
                    End If
                Next
            Catch ex As Exception
                Log("ADVERTENCIA recolectando docs: " & ex.Message)
            End Try

            If docsToClose.Count = 0 Then Exit For

            For Each doc As Document In docsToClose
                Try
                    Dim fileName As String = ""
                    Try : fileName = doc.FullFileName : Catch : End Try

                    doc.Close(True)
                    closedCount += 1
                    Log("Cerrado: " & IOPath.GetFileName(fileName))
                Catch ex As Exception
                    Log("No se pudo cerrar documento: " & ex.Message)
                End Try
            Next

            If docsToClose.Count = 0 Then Exit For
            System.Threading.Thread.Sleep(500)
        Next

        Log("Total documentos cerrados: " & closedCount)
    End Sub

    Private Sub CopyDirectory(sourceDir As String, targetDir As String)
        If sourceDir.ToLower().Contains("\oldversions") Then Exit Sub

        IODirectory.CreateDirectory(targetDir)

        For Each filePath As String In IODirectory.GetFiles(sourceDir)
            IOFile.Copy(filePath, IOPath.Combine(targetDir, IOPath.GetFileName(filePath)), True)
        Next

        For Each subFolder As String In IODirectory.GetDirectories(sourceDir)
            CopyDirectory(subFolder, IOPath.Combine(targetDir, IOPath.GetFileName(subFolder)))
        Next
    End Sub

    Private Sub ActivateProject(ipjPath As String)
        Dim mgr As DesignProjectManager = _app.DesignProjectManager
        Dim proj As DesignProject = mgr.DesignProjects.AddExisting(ipjPath)
        proj.Activate()
    End Sub

    Private Function BuildProjectName(data As ProjectData) As String
        Return CleanName(data.Modelo & " - " & data.Maquina & " - " & data.OT & " - " & data.Cliente)
    End Function

    Private Function CleanName(value As String) As String
        Dim cleanValue As String = value
        For Each c As Char In IOPath.GetInvalidFileNameChars()
            cleanValue = cleanValue.Replace(c, "-"c)
        Next
        Return cleanValue.Trim()
    End Function

    Private Function NormalizeCode(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""
        Dim parts() As String = value.Split(ChrW(46))
        If parts.Length >= 2 Then
            Return parts(0) & "." & parts(1)
        End If
        Return value.Trim()
    End Function

End Class

Public Class ProjectData
    Public Property Modelo As String
    Public Property Maquina As String
    Public Property OT As String
    Public Property Cliente As String
End Class
