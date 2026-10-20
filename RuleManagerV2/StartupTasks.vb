Imports Inventor
Imports IO = System.IO
Imports System.Windows.Forms

Public Class StartupTasks

    Private Shared _invApp As Inventor.Application
    Private Shared _appEvents As ApplicationEvents
    Private Shared _userInputEvents As UserInputEvents
    Private Shared _timer As System.Windows.Forms.Timer
    Private Shared _ultimoProyecto As String = ""
    Private Shared _proyectoYaOrganizado As String = ""
    Private Shared _tickCount As Integer = 0
    Private Shared ReadOnly _piezasProcesadas As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    Private Const LOG_PATH As String = "C:\Temp\Reorder.log"

    ' Lista blanca de librerías
    Private Shared ReadOnly _libreriasPermitidas As New List(Of String) From {
        "METALES DB",
        "PLASTICOS DB",
        "DE BLASI MAT"
    }

    ' Protegidas del sistema
    Private Shared ReadOnly _libreriasProtegidas As New List(Of String) From {
        "PhysicalMaterial",
        "InventorMaterialLibrary"
    }

    Private Const CARPETA_LIBRERIAS As String = "D:\Bibliotecas\Materiales"

    ' =========================================================
    ' ESTRUCTURA DE CARPETAS DEL PROYECTO
    ' =========================================================
    Private Shared ReadOnly _carpetasProyecto As New List(Of String) From {
       "01 - DIBUJOS",
        "02 - PLANOS",
        "03 - CORTES",
        "04 - COMERCIALES",
        "05 - PIEZAS COMUNES",
        "06 - MULTIMEDIA",
        "07 - AUTOCAD - SAT",
        "08 - MAQUINAS DE REFERENCIAS"
    }

    ' =========================================================
    ' LOG FÍSICO
    ' =========================================================
    Private Shared Sub WriteLog(mensaje As String)
        Try
            Dim dir As String = IO.Path.GetDirectoryName(LOG_PATH)
            If Not IO.Directory.Exists(dir) Then
                IO.Directory.CreateDirectory(dir)
            End If

            If IO.File.Exists(LOG_PATH) Then
                Dim fi As New IO.FileInfo(LOG_PATH)
                If fi.IsReadOnly Then fi.IsReadOnly = False
            End If

            Using sw As New IO.StreamWriter(LOG_PATH, True)
                sw.WriteLine($"{Now:yyyy-MM-dd HH:mm:ss} | [StartupTasks] {mensaje}")
            End Using
        Catch
        End Try
    End Sub

    Private Shared Sub DebugPrint(mensaje As String)
        System.Diagnostics.Debug.WriteLine("[RuleManager] " & mensaje)
        WriteLog(mensaje)
    End Sub

    Public Shared Sub Initialize(invApp As Inventor.Application)
        _invApp = invApp
        _appEvents = _invApp.ApplicationEvents
        AddHandler _appEvents.OnActiveProjectChanged, AddressOf OnActiveProjectChanged
        AddHandler _appEvents.OnNewDocument, AddressOf OnNewDocument
        AddHandler _appEvents.OnNewEditObject, AddressOf OnNewEditObject

        Try
            _userInputEvents = _invApp.CommandManager.UserInputEvents
            AddHandler _userInputEvents.OnTerminateCommand, AddressOf OnCommandTerminate
        Catch ex As Exception
            DebugPrint("No se pudo enlazar UserInputEvents: " & ex.Message)
        End Try

        _timer = New System.Windows.Forms.Timer()
        _timer.Interval = 2000
        AddHandler _timer.Tick, AddressOf OnMonitorTick
        _timer.Start()

        DebugPrint("StartupTasks inicializado con OnNewDocument, OnNewEditObject y UserInputEvents")
    End Sub

    Public Shared Sub Shutdown()
        Try
            If _appEvents IsNot Nothing Then
                RemoveHandler _appEvents.OnActiveProjectChanged, AddressOf OnActiveProjectChanged
                RemoveHandler _appEvents.OnNewDocument, AddressOf OnNewDocument
                RemoveHandler _appEvents.OnNewEditObject, AddressOf OnNewEditObject
            End If
            If _userInputEvents IsNot Nothing Then
                RemoveHandler _userInputEvents.OnTerminateCommand, AddressOf OnCommandTerminate
            End If
            If _timer IsNot Nothing Then
                _timer.Stop()
                _timer.Dispose()
            End If
        Catch
        End Try
    End Sub

    ' =========================================================
    ' EVENTO 1: CREACIÓN DE NUEVO DOCUMENTO (DOC DIRECTO)
    ' =========================================================
    Private Shared Sub OnNewDocument(
        ByVal DocumentObject As Inventor._Document,
        ByVal BeforeOrAfter As EventTimingEnum,
        ByVal Context As NameValueMap,
        ByRef HandlingCode As HandlingCodeEnum)

        If BeforeOrAfter <> EventTimingEnum.kAfter Then Exit Sub
        If DocumentObject Is Nothing Then Exit Sub
        If DocumentObject.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then Exit Sub

        VerificarYEjecutarSeleccionBaseParaPieza(DocumentObject, "OnNewDocument", Context, False)
    End Sub

    ' =========================================================
    ' EVENTO 2: ENTRADA EN EDICIÓN DE OBJETO (IN SITU EN ENSAMBLAJE)
    ' =========================================================
    Private Shared Sub OnNewEditObject(
        ByVal EditObject As Object,
        ByVal BeforeOrAfter As EventTimingEnum,
        ByVal Context As NameValueMap,
        ByRef HandlingCode As HandlingCodeEnum)

        If BeforeOrAfter <> EventTimingEnum.kAfter Then Exit Sub
        If Not AddinSettings.Current.ActivarSeleccionBaseEnPlantilla Then Exit Sub

        Try
            Dim partDoc As PartDocument = Nothing

            ' Caso A: Documento activo en edición in situ
            If _invApp.ActiveEditDocument IsNot Nothing AndAlso
               _invApp.ActiveEditDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                partDoc = DirectCast(_invApp.ActiveEditDocument, PartDocument)

                ' Caso B: EditObject es ComponentOccurrence o PartComponentDefinition
            ElseIf TypeOf EditObject Is ComponentOccurrence Then
                Dim occ As ComponentOccurrence = DirectCast(EditObject, ComponentOccurrence)
                If occ.Definition.Document.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    partDoc = DirectCast(occ.Definition.Document, PartDocument)
                End If
            ElseIf TypeOf EditObject Is PartComponentDefinition Then
                Dim def As PartComponentDefinition = DirectCast(EditObject, PartComponentDefinition)
                partDoc = DirectCast(def.Document, PartDocument)
            End If

            If partDoc IsNot Nothing Then
                ' Si estamos dentro de un ensamblaje, es una edición in situ
                Dim enEnsamblaje As Boolean = (_invApp.ActiveDocument IsNot Nothing AndAlso
                                               _invApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject)
                VerificarYEjecutarSeleccionBaseParaPieza(DirectCast(partDoc, Document), "OnNewEditObject (In Situ)", Context, enEnsamblaje)
            End If

        Catch ex As Exception
            DebugPrint("Excepción en OnNewEditObject: " & ex.Message)
        End Try
    End Sub

    ' =========================================================
    ' EVENTO 3: TERMINACIÓN DE COMANDO NATIVO DE CREACIÓN IN SITU
    ' =========================================================
    Private Shared Sub OnCommandTerminate(
        ByVal CommandName As String,
        ByVal Context As NameValueMap)

        If Not AddinSettings.Current.ActivarSeleccionBaseEnPlantilla Then Exit Sub

        Try
            If String.IsNullOrEmpty(CommandName) Then Exit Sub

            ' Comandos nativos de Inventor de creación in situ dentro de ensamblaje
            Dim cmdUpper As String = CommandName.ToUpper()
            If cmdUpper.Contains("ASSEMBLYCREATECOMPONENT") OrElse
               cmdUpper.Contains("ASSEMBLYCREATEINPLACE") OrElse
               cmdUpper = "ASSEMBLYCREATECOMPONENTCMD" Then

                DebugPrint("Comando creación in situ detectado: " & CommandName)

                ' En este instante Inventor ya colocó la nueva pieza y está activa en ActiveEditDocument
                If _invApp.ActiveEditDocument IsNot Nothing AndAlso
                   _invApp.ActiveEditDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then

                    VerificarYEjecutarSeleccionBaseParaPieza(_invApp.ActiveEditDocument, "OnCommandTerminate (In Situ)", Context, True)
                End If
            End If
        Catch ex As Exception
            DebugPrint("Excepción en OnCommandTerminate: " & ex.Message)
        End Try
    End Sub

    ' =========================================================
    ' LÓGICA CENTRAL DE DETECCIÓN Y DISPARO DE SELECCIÓN BASE
    ' =========================================================
    Private Shared Sub VerificarYEjecutarSeleccionBaseParaPieza(
        doc As Document,
        origen As String,
        context As NameValueMap,
        esInSituEnEnsamblaje As Boolean)

        If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then Exit Sub
        If Not AddinSettings.Current.ActivarSeleccionBaseEnPlantilla Then Exit Sub

        Try
            Dim docId As String = ""
            Try
                docId = doc.InternalName
            Catch
            End Try
            If String.IsNullOrEmpty(docId) Then docId = doc.FullFileName
            If String.IsNullOrEmpty(docId) Then docId = doc.DisplayName

            ' Evitar procesar dos veces el mismo documento en la sesión
            If Not String.IsNullOrEmpty(docId) AndAlso _piezasProcesadas.Contains(docId) Then
                Exit Sub
            End If

            ' Si la pieza ya tiene la iProperty MATERIAL_BASE con contenido, ya fue configurada anteriormente
            Try
                Dim userProps As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
                If userProps IsNot Nothing Then
                    For Each p As Inventor.Property In userProps
                        If p.Name.Equals("MATERIAL_BASE", StringComparison.OrdinalIgnoreCase) AndAlso
                           p.Value IsNot Nothing AndAlso
                           Not String.IsNullOrWhiteSpace(p.Value.ToString()) Then
                            If Not String.IsNullOrEmpty(docId) Then _piezasProcesadas.Add(docId)
                            Exit Sub
                        End If
                    Next
                End If
            Catch
            End Try

            Dim esPlantillaObjetivo As Boolean = False
            Dim plantillaBuscada As String = AddinSettings.Current.PlantillaPiezaDefault.ToUpper()

            ' 1. Verificar Context si contiene el nombre de la plantilla de forma segura
            If context IsNot Nothing Then
                For i As Integer = 1 To context.Count
                    Try
                        Dim kName As String = context.Name(i)
                        Dim valObj As Object = context.Value(kName)
                        If valObj IsNot Nothing Then
                            Dim valStr As String = valObj.ToString().ToUpper()
                            If valStr.Contains("PIEZA DB 2") OrElse valStr.Contains(plantillaBuscada) Then
                                esPlantillaObjetivo = True
                                Exit For
                            End If
                        End If
                    Catch
                    End Try
                Next
            End If

            ' 2. Verificar nombre o ruta del documento
            If Not esPlantillaObjetivo Then
                Dim full As String = (doc.FullFileName & " " & doc.DisplayName).ToUpper()
                If full.Contains("PIEZA DB 2") OrElse full.Contains(plantillaBuscada) Then
                    esPlantillaObjetivo = True
                End If
            End If

            ' 3. Verificar propiedades de usuario o parámetros importados de la plantilla
            If Not esPlantillaObjetivo Then
                Try
                    Dim userProps As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
                    If userProps IsNot Nothing Then
                        For Each p As Inventor.Property In userProps
                            Dim pName As String = p.Name.ToUpper()
                            Dim pVal As String = If(p.Value IsNot Nothing, p.Value.ToString().ToUpper(), "")
                            If pName.Contains("PIEZA DB 2") OrElse pVal.Contains("PIEZA DB 2") OrElse
                               (pName = "PLANTILLA" AndAlso pVal.Contains("PIEZA DB")) Then
                                esPlantillaObjetivo = True
                                Exit For
                            End If
                        Next
                    End If
                Catch
                End Try
            End If

            ' 4. Si proviene de una creación in situ dentro de ensamblaje:
            ' En la empresa la plantilla predeterminada de nuevas piezas es siempre PIEZA DB 2.
            ' Si la pieza no tiene operaciones modeladas aún (Features.Count = 0) y no tiene MATERIAL_BASE,
            ' es una pieza recién creada in situ con la plantilla objetivo.
            If Not esPlantillaObjetivo AndAlso esInSituEnEnsamblaje Then
                Try
                    Dim partDoc As PartDocument = DirectCast(doc, PartDocument)
                    If partDoc.ComponentDefinition.Features.Count = 0 Then
                        esPlantillaObjetivo = True
                    End If
                Catch
                End Try
            End If

            If esPlantillaObjetivo Then
                If Not String.IsNullOrEmpty(docId) Then _piezasProcesadas.Add(docId)
                DebugPrint($"Pieza detectada para Selección Base ({origen}) [{doc.DisplayName}] -> Lanzando diálogo...")

                ' Disparar mediante Timer para permitir que Inventor estabilice el viewport y el foco de edición
                Dim delayTimer As New System.Windows.Forms.Timer()
                delayTimer.Interval = 350
                AddHandler delayTimer.Tick, Sub(sender As Object, ev As EventArgs)
                                                delayTimer.Stop()
                                                delayTimer.Dispose()
                                                Try
                                                    Dim engine As New SeleccionBaseEngine(_invApp)
                                                    engine.Ejecutar(doc)
                                                Catch ex As Exception
                                                    DebugPrint("Error al ejecutar Selección Base: " & ex.Message)
                                                End Try
                                            End Sub
                delayTimer.Start()
            End If

        Catch ex As Exception
            DebugPrint("Excepción en VerificarYEjecutarSeleccionBaseParaPieza: " & ex.Message)
        End Try
    End Sub

    Public Shared Sub LoadLibrariesForCurrentProject()
        DebugPrint("Carga forzada solicitada")
        SincronizarLibrerias()
        CrearEstructuraCarpetas()
    End Sub

    Private Shared Sub OnActiveProjectChanged(
        ByVal ProjectObject As DesignProject,
        ByVal BeforeOrAfter As EventTimingEnum,
        ByVal Context As NameValueMap,
        ByRef HandlingCode As HandlingCodeEnum)

        If BeforeOrAfter <> EventTimingEnum.kAfter Then Exit Sub

        If ProjectObject.FullFileName <> _ultimoProyecto Then
            _proyectoYaOrganizado = ""
        End If

        DebugPrint("OnActiveProjectChanged: " & ProjectObject.FullFileName)
        _ultimoProyecto = ProjectObject.FullFileName
        SincronizarLibrerias()
        CrearEstructuraCarpetas()
    End Sub

    Private Shared Sub OnMonitorTick(sender As Object, e As EventArgs)
        Try
            _tickCount += 1
            If _tickCount < 10 Then
                Return
            End If
            _tickCount = 0

            Dim proj As DesignProject = _invApp.DesignProjectManager.ActiveDesignProject
            If proj Is Nothing Then Exit Sub

            If proj.FullFileName = _ultimoProyecto Then
                Dim libs As ProjectAssetLibraries = proj.MaterialLibraries
                Dim necesitaSync As Boolean = False

                Dim i As Integer
                For i = 0 To _libreriasPermitidas.Count - 1
                    Dim rutaPermitida As String = IO.Path.Combine(CARPETA_LIBRERIAS, _libreriasPermitidas(i) & ".adsklib")
                    If Not IsLibraryLoaded(libs, rutaPermitida) Then
                        necesitaSync = True
                        Exit For
                    End If
                Next

                If necesitaSync Then
                    DebugPrint("Timer detecto que falta libreria, sincronizando...")
                    SincronizarLibrerias()
                End If

                Exit Sub
            End If

            _ultimoProyecto = proj.FullFileName
            _proyectoYaOrganizado = ""
            DebugPrint("Timer detecto cambio de proyecto: " & proj.FullFileName)
            SincronizarLibrerias()
            CrearEstructuraCarpetas()

        Catch
        End Try
    End Sub

    ' =========================================================
    ' SINCRONIZAR LIBRERIAS
    ' =========================================================
    Private Shared Sub SincronizarLibrerias()
        Try
            Dim proj As DesignProject = _invApp.DesignProjectManager.ActiveDesignProject
            If proj Is Nothing Then Exit Sub

            DebugPrint("=== SINCRONIZANDO LIBRERIAS ===")
            DebugPrint("Proyecto: " & proj.FullFileName)

            Try
                proj.Save()
                DebugPrint("Proyecto guardado")
            Catch ex As Exception
                DebugPrint("No se pudo guardar proyecto: " & ex.Message)
            End Try

            Dim libs As ProjectAssetLibraries = proj.MaterialLibraries

            Dim k As Integer
            For k = libs.Count To 1 Step -1
                Try
                    Dim al As Object = libs(k)
                    Dim nombre As String = GetNameSafe(al)
                    If nombre = "" Then Continue For

                    Dim esPermitida As Boolean = False
                    Dim esProtegida As Boolean = False
                    Dim j As Integer

                    For j = 0 To _libreriasPermitidas.Count - 1
                        If _libreriasPermitidas(j) = nombre Then
                            esPermitida = True
                            Exit For
                        End If
                    Next

                    For j = 0 To _libreriasProtegidas.Count - 1
                        If _libreriasProtegidas(j) = nombre Then
                            esProtegida = True
                            Exit For
                        End If
                    Next

                    If Not esPermitida And Not esProtegida Then
                        DebugPrint("  -> Eliminando: " & nombre)
                        Try
                            al.Delete()
                            DebugPrint("     Eliminada OK")
                        Catch ex As Exception
                            DebugPrint("     No se pudo eliminar: " & ex.Message)
                        End Try
                    End If
                Catch
                End Try
            Next

            Dim cargadas As Integer = 0
            Dim omitidas As Integer = 0

            For Each nombreLib As String In _libreriasPermitidas
                Dim rutaLib As String = IO.Path.Combine(CARPETA_LIBRERIAS, nombreLib & ".adsklib")

                If Not IO.File.Exists(rutaLib) Then
                    DebugPrint("No existe en disco: " & rutaLib)
                    Continue For
                End If

                If IsLibraryLoaded(libs, rutaLib) Then
                    omitidas = omitidas + 1
                    DebugPrint("Ya cargada: " & nombreLib)
                    Continue For
                End If

                DebugPrint("Intentando cargar: " & nombreLib)
                Try
                    libs.Add(rutaLib)
                    cargadas = cargadas + 1
                    DebugPrint("  -> OK")
                Catch ex As ArgumentException
                    omitidas = omitidas + 1
                    DebugPrint("  -> E_INVALIDARG")
                Catch ex As Runtime.InteropServices.COMException
                    DebugPrint("  -> E_FAIL")
                Catch ex As Exception
                    DebugPrint("  -> ERROR: " & ex.Message)
                End Try
            Next

            If cargadas > 0 Then
                Try
                    proj.Save()
                    MessageBox.Show("Se cargaron " & cargadas.ToString() & " libreria(s).", "Librerias sincronizadas", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch
                End Try
            End If

            DebugPrint("=== RESUMEN LIBRERIAS: Cargadas=" & cargadas.ToString() & " Omitidas=" & omitidas.ToString() & " ===")

        Catch ex As Exception
            DebugPrint("ERROR GLOBAL LIBRERIAS: " & ex.Message)
        End Try
    End Sub

    ' =========================================================
    ' CREAR ESTRUCTURA DE CARPETAS
    ' =========================================================
    Private Shared Sub CrearEstructuraCarpetas()

        Try
            ' === NUEVO: RESPETAR OPCION DEL USUARIO ===
            If Not AddinSettings.Current.CrearCarpetasProyecto Then
                DebugPrint("Creacion de carpetas deshabilitada en OPCIONES. Se omite.")
                Exit Sub
            End If

            Dim proj As DesignProject = _invApp.DesignProjectManager.ActiveDesignProject
            If proj Is Nothing Then
                DebugPrint("No hay proyecto activo para crear carpetas")
                Exit Sub
            End If

            Dim rutaProyecto As String = IO.Path.GetDirectoryName(proj.FullFileName)
            If Not IO.Directory.Exists(rutaProyecto) Then
                DebugPrint("La carpeta del proyecto no existe: " & rutaProyecto)
                Exit Sub
            End If

            DebugPrint("=== CREANDO CARPETAS ===")
            DebugPrint("Ruta base: " & rutaProyecto)

            Dim creadas As Integer = 0
            Dim existentes As Integer = 0

            For Each nombreCarpeta As String In _carpetasProyecto
                Dim rutaCarpeta As String = IO.Path.Combine(rutaProyecto, nombreCarpeta)

                If IO.Directory.Exists(rutaCarpeta) Then
                    existentes = existentes + 1
                    DebugPrint("Ya existe: " & nombreCarpeta)
                Else
                    Try
                        IO.Directory.CreateDirectory(rutaCarpeta)
                        creadas = creadas + 1
                        DebugPrint("Creada: " & nombreCarpeta)
                    Catch ex As Exception
                        DebugPrint("ERROR al crear " & nombreCarpeta & ": " & ex.Message)
                    End Try
                End If
            Next

            DebugPrint("=== RESUMEN CARPETAS: Creadas=" & creadas.ToString() & " Existentes=" & existentes.ToString() & " ===")

            If creadas > 0 Then
                MessageBox.Show("Se crearon " & creadas.ToString() & " carpeta(s) en el proyecto.", "Carpetas creadas", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            DebugPrint("ERROR GLOBAL CARPETAS: " & ex.Message)
        End Try
    End Sub

    ' =========================================================
    ' ORGANIZAR ARCHIVOS DEL PROYECTO - CON REVISIÓN DE USUARIO
    ' =========================================================
    Public Shared Sub OrganizarArchivosProyecto(inventor As Inventor.Application)
        Dim proj = inventor.DesignProjectManager.ActiveDesignProject
        If proj Is Nothing Then Exit Sub

        If proj.FullFileName = _proyectoYaOrganizado Then
            DebugPrint("Organizacion ya realizada para este proyecto en esta sesion. Saltando.")
            MessageBox.Show(
                "Este proyecto ya fue organizado en esta sesión." & vbCrLf &
                "Reinicia Inventor si necesitas reorganizar.",
                "Organizar Proyecto",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim rootPath = IO.Path.GetDirectoryName(proj.FullFileName)

        DebugPrint("=== INICIO ORGANIZAR ARCHIVOS ===")
        DebugPrint("Ruta raiz: " & rootPath)

        ' --- FASE 0: PLANIFICAR TODOS LOS MOVIMIENTOS ---
        Dim allFiles = IO.Directory.GetFiles(rootPath, "*.*", IO.SearchOption.AllDirectories)
        Dim plannedMoves As New List(Of FileMoveItem)

        For Each file In allFiles
            Dim item As New FileMoveItem(file, rootPath)
            Dim destino As String = FileClassifier.GetDestinationPath(inventor, file, rootPath, rootPath)

            If String.IsNullOrEmpty(destino) Then
                Dim targetFolder As String = FileClassifier.GetTargetFolder(inventor, file, rootPath)
                If targetFolder = "__UNCLASSIFIED__" Then
                    item.IsClassified = False
                    item.FinalDestination = ""
                    item.TargetFolder = ""
                Else
                    item.IsExcluded = True
                    item.IsClassified = True
                End If
            ElseIf destino = "__UNCLASSIFIED__" Then
                item.IsClassified = False
                item.FinalDestination = ""
                item.TargetFolder = ""
            Else
                item.IsClassified = True
                item.FinalDestination = destino
                item.TargetFolder = destino.Substring(rootPath.Length).TrimStart("\"c).Split("\"c)(0)
            End If

            plannedMoves.Add(item)
        Next

        ' --- FASE 1: MOSTRAR DIÁLOGO DE REVISIÓN ---
        Dim hwnd As IntPtr = IntPtr.Zero
        Try
            hwnd = New IntPtr(inventor.MainFrameHWND)
        Catch
        End Try

        Dim owner As IWin32Window = Nothing
        If hwnd <> IntPtr.Zero Then
            owner = New WindowWrapper(hwnd)
        End If

        Using reviewDlg As New ReviewDialog(plannedMoves, rootPath, inventor)
            Dim result As DialogResult = reviewDlg.ShowDialog(owner)

            If result = DialogResult.OK Then
                _proyectoYaOrganizado = proj.FullFileName
                DebugPrint("=== FIN ORGANIZAR ARCHIVOS (OK) ===")
            Else
                DebugPrint("Organización cancelada por el usuario.")
                DebugPrint("=== FIN ORGANIZAR ARCHIVOS (CANCELADO) ===")
            End If
        End Using
    End Sub

    ' =========================================================
    ' HELPERS
    ' =========================================================
    Private Shared Function IsLibraryLoaded(libs As ProjectAssetLibraries, rutaLib As String) As Boolean
        Dim targetName As String = IO.Path.GetFileNameWithoutExtension(rutaLib).ToLower()

        Dim i As Integer
        For i = 1 To libs.Count
            Try
                Dim al As Object = libs(i)
                Dim nombre As String = GetNameSafe(al)
                If nombre.ToLower() = targetName Then
                    Return True
                End If
            Catch
            End Try
        Next
        Return False
    End Function

    Private Shared Function GetNameSafe(al As Object) As String
        Dim nombre As String = ""

        Try
            nombre = CStr(CallByName(al, "DisplayName", CallType.Get))
        Catch
            nombre = ""
        End Try

        If nombre = "" Then
            Try
                nombre = CStr(CallByName(al, "Name", CallType.Get))
            Catch
                nombre = ""
            End Try
        End If

        If nombre = "" Then
            Try
                nombre = CStr(CallByName(al, "InternalName", CallType.Get))
            Catch
                nombre = ""
            End Try
        End If

        Return nombre
    End Function

End Class