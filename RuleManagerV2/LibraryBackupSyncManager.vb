Imports System
Imports System.Collections.Concurrent
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Security.AccessControl
Imports System.Security.Principal
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Inventor
' Alias explícito para evitar colisión de nombres entre System e Inventor (Errores BC30561 y BC30518)
Imports SysEnv = System.Environment
Imports Path = System.IO.Path
Imports File = System.IO.File
Imports Directory = System.IO.Directory
Imports FileInfo = System.IO.FileInfo

''' <summary>
''' Estado de sincronización de un archivo individual entre el equipo local y el servidor de red.
''' </summary>
Public Enum FileSyncStatusEnum
    Identical           ' Mismo tamaño y fecha similar
    ServerNewer         ' El servidor tiene una versión más reciente (requiere PULL)
    LocalNewer          ' El equipo local tiene una versión más reciente (requiere PUSH)
    OnlyOnServer        ' Existe en el servidor pero no en el equipo local
    OnlyOnLocal         ' Existe en el equipo local pero no en el servidor
    Conflict            ' Ambos modificados recientemente o con discrepancias
End Enum

''' <summary>
''' Intervalo programado para la comprobación y backup de librerías.
''' </summary>
Public Enum IntervaloBackupEnum
    Manual = 0
    Diario = 1
    Semanal = 2
    Mensual = 3
End Enum

''' <summary>
''' Estructura liviana en memoria para metadatos rápidos de archivos
''' extraídos de la cabecera SMB (sin llamadas RPC adicionales de seguridad).
''' </summary>
Public Class FileQuickMeta
    Public Property RelativePath As String
    Public Property FullPath As String
    Public Property Size As Long
    Public Property LastWriteUtc As DateTime
End Class

''' <summary>
''' Representa el resultado comparativo de un archivo entre local y servidor.
''' </summary>
Public Class FileComparisonItem
    Public Property RelativePath As String
    Public Property LocalFullPath As String
    Public Property ServerFullPath As String
    Public Property LocalSize As Long
    Public Property ServerSize As Long
    Public Property LocalModified As DateTime
    Public Property ServerModified As DateTime
    Public Property LocalModifiedBy As String = "-"
    Public Property ServerModifiedBy As String = "-"
    Public Property Status As FileSyncStatusEnum
    Public Property InternalDetails As New List(Of String)

    Public ReadOnly Property StatusDescription As String
        Get
            Select Case Status
                Case FileSyncStatusEnum.Identical
                    Return "✔ Sincronizado (Idéntico)"
                Case FileSyncStatusEnum.ServerNewer
                    Return "⬇ Servidor más reciente (PULL)"
                Case FileSyncStatusEnum.LocalNewer
                    Return "⬆ Local más reciente (PUSH)"
                Case FileSyncStatusEnum.OnlyOnServer
                    Return "⬇ Solo en servidor (PULL requerido)"
                Case FileSyncStatusEnum.OnlyOnLocal
                    Return "⬆ Solo local (PUSH pendiente)"
                Case Else
                    Return "⚠ Conflicto / Desconocido"
            End Select
        End Get
    End Property
End Class

''' <summary>
''' Motor de sincronización y backup para bibliotecas CAD de Autodesk Inventor.
''' Optimizado para conexiones de alta velocidad (5GbE) con escaneo paralelo,
''' transferencias por búferes ampliados e inspección profunda de materiales (.adsklib / .asdklib).
''' </summary>
Public Class LibraryBackupSyncManager

    Public Const DEFAULT_LOCAL_PATH As String = "D:\Bibliotecas"
    Public Const DEFAULT_SERVER_PATH As String = "\\192.168.0.229\Dpto_Tec\COMERCIALES INVENTOR\DATOS DE DISEÑO INVENTOR\Bibliotecas"

    Private _invApp As Inventor.Application
    Private ReadOnly _cachePropietarios As New ConcurrentDictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

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
    End Sub

    Public Function ObtenerInventorApp() As Inventor.Application
        If _invApp IsNot Nothing Then Return _invApp
        Try
            _invApp = DirectCast(Microsoft.VisualBasic.Interaction.GetObject(Nothing, "Inventor.Application"), Inventor.Application)
        Catch
        End Try
        Return _invApp
    End Function

    ''' <summary>
    ''' Determina si según el intervalo configurado y la fecha del último sincronizado corresponde ejecutar sincronización automática.
    ''' </summary>
    Public Function CorrespondeSincronizarAutomaticamente(intervalo As IntervaloBackupEnum, ultimaSinc As DateTime) As Boolean
        If intervalo = IntervaloBackupEnum.Manual Then Return False
        If ultimaSinc = DateTime.MinValue Then Return True

        Dim tiempoTranscurrido = DateTime.Now - ultimaSinc

        Select Case intervalo
            Case IntervaloBackupEnum.Diario
                Return tiempoTranscurrido.TotalHours >= 20.0 OrElse ultimaSinc.Date < DateTime.Today
            Case IntervaloBackupEnum.Semanal
                Return tiempoTranscurrido.TotalDays >= 7.0
            Case IntervaloBackupEnum.Mensual
                Return tiempoTranscurrido.TotalDays >= 30.0
            Case Else
                Return False
        End Select
    End Function

    ''' <summary>
    ''' Comprueba si un archivo o directorio debe ser excluido de la sincronización.
    ''' Ignora carpetas de logs ("\Sync_Logs\", "\Logs\") y archivos de log temporales.
    ''' </summary>
    Public Shared Function EsArchivoIgnorado(rutaCompleta As String) As Boolean
        If String.IsNullOrWhiteSpace(rutaCompleta) Then Return True

        Dim rutaNorm = rutaCompleta.Replace("/"c, "\"c)

        ' Ignorar carpetas de logs internas
        If rutaNorm.IndexOf("\Sync_Logs\", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
           rutaNorm.IndexOf("\Logs\", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
           rutaNorm.IndexOf("\Log\", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
           rutaNorm.EndsWith("\Sync_Logs", StringComparison.OrdinalIgnoreCase) OrElse
           rutaNorm.EndsWith("\Logs", StringComparison.OrdinalIgnoreCase) Then
            Return True
        End If

        ' Ignorar extensiones de logs y archivos temporales
        Dim ext = Path.GetExtension(rutaNorm).ToLowerInvariant()
        If ext = ".log" OrElse ext = ".tmp" OrElse ext = ".bak" OrElse ext = ".lck" OrElse ext = ".old" Then
            Return True
        End If

        ' Ignorar nombres específicos de logs de sincronización
        Dim nombreArchivo = Path.GetFileName(rutaNorm).ToLowerInvariant()
        If nombreArchivo.StartsWith("sync_log_") OrElse
           nombreArchivo = "changelog_general.txt" OrElse
           nombreArchivo = "sync_log.txt" OrElse
           nombreArchivo.StartsWith("backup_log_") Then
            Return True
        End If

        Return False
    End Function

    ''' <summary>
    ''' Compara archivo a archivo entre la carpeta local y el servidor de red.
    ''' Ejecuta el escaneo de directorios local y remoto en paralelo para saturar la red 5GbE.
    ''' </summary>
    Public Function CompararArchivos(localRoot As String, serverRoot As String, Optional progreso As IProgress(Of String) = Nothing) As List(Of FileComparisonItem)
        Dim resultados As New List(Of FileComparisonItem)

        If Not Directory.Exists(localRoot) AndAlso Not Directory.Exists(serverRoot) Then
            Return resultados
        End If

        progreso?.Report("Escaneando carpetas local y servidor en paralelo a alta velocidad...")

        ' Escaneo concurrente de árbol local y árbol en red SMB (hilos independientes)
        Dim taskLocal = Task.Run(Function() EscanearDirectorioRapido(localRoot))
        Dim taskServer = Task.Run(Function() EscanearDirectorioRapido(serverRoot))

        Task.WaitAll(taskLocal, taskServer)

        Dim diccLocal = taskLocal.Result
        Dim diccServer = taskServer.Result

        progreso?.Report($"Indexación completada: {diccLocal.Count} archivos locales, {diccServer.Count} en servidor. Comparando metadatos...")

        ' Unificar conjunto de todas las rutas relativas
        Dim todasLasRutas As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each k In diccLocal.Keys : todasLasRutas.Add(k) : Next
        For Each k In diccServer.Keys : todasLasRutas.Add(k) : Next

        Dim itemsModificadosParaInspeccionar As New List(Of FileComparisonItem)

        For Each relPath In todasLasRutas
            Dim item As New FileComparisonItem With {
                .RelativePath = relPath,
                .LocalFullPath = System.IO.Path.Combine(localRoot, relPath),
                .ServerFullPath = System.IO.Path.Combine(serverRoot, relPath)
            }

            Dim tieneLocal = diccLocal.ContainsKey(relPath)
            Dim tieneServer = diccServer.ContainsKey(relPath)

            If tieneLocal AndAlso Not tieneServer Then
                Dim loc = diccLocal(relPath)
                item.LocalSize = loc.Size
                item.LocalModified = loc.LastWriteUtc.ToLocalTime()
                item.Status = FileSyncStatusEnum.OnlyOnLocal
                item.InternalDetails.Add("Archivo nuevo en el equipo local (pendiente de subir PUSH al servidor).")

            ElseIf tieneServer AndAlso Not tieneLocal Then
                Dim srv = diccServer(relPath)
                item.ServerSize = srv.Size
                item.ServerModified = srv.LastWriteUtc.ToLocalTime()
                item.Status = FileSyncStatusEnum.OnlyOnServer
                item.InternalDetails.Add("Archivo existente únicamente en el servidor (requiere descarga PULL).")

            Else
                ' Existe en ambos
                Dim loc = diccLocal(relPath)
                Dim srv = diccServer(relPath)

                item.LocalSize = loc.Size
                item.ServerSize = srv.Size
                item.LocalModified = loc.LastWriteUtc.ToLocalTime()
                item.ServerModified = srv.LastWriteUtc.ToLocalTime()

                ' Tolerancia de 2 segundos por diferencias entre sistemas de archivos FAT y NTFS
                Dim difSegundos = (srv.LastWriteUtc - loc.LastWriteUtc).TotalSeconds

                If Math.Abs(difSegundos) <= 2 AndAlso loc.Size = srv.Size Then
                    item.Status = FileSyncStatusEnum.Identical
                ElseIf srv.LastWriteUtc > loc.LastWriteUtc Then
                    item.Status = FileSyncStatusEnum.ServerNewer
                    itemsModificadosParaInspeccionar.Add(item)
                Else
                    item.Status = FileSyncStatusEnum.LocalNewer
                    itemsModificadosParaInspeccionar.Add(item)
                End If
            End If

            resultados.Add(item)
        Next

        ' Inspección interna paralela de archivos modificados (aprovechando ancho de banda y multinúcleo)
        If itemsModificadosParaInspeccionar.Count > 0 Then
            progreso?.Report($"Inspeccionando contenido interno de {itemsModificadosParaInspeccionar.Count} archivo(s) modificados en paralelo...")

            Dim maxHilos = Math.Min(8, Math.Max(2, SysEnv.ProcessorCount))
            Dim parOpts As New ParallelOptions With {.MaxDegreeOfParallelism = maxHilos}

            Parallel.ForEach(itemsModificadosParaInspeccionar, parOpts,
                Sub(modItem As FileComparisonItem)
                    Try
                        If modItem.Status = FileSyncStatusEnum.ServerNewer Then
                            ' El servidor es más reciente: archivoAnterior = Local, archivoNuevo = Servidor
                            modItem.InternalDetails = InspeccionarDiferenciasInternas(modItem.LocalFullPath, modItem.ServerFullPath)
                        ElseIf modItem.Status = FileSyncStatusEnum.LocalNewer Then
                            ' El local es más reciente: archivoAnterior = Servidor, archivoNuevo = Local
                            modItem.InternalDetails = InspeccionarDiferenciasInternas(modItem.ServerFullPath, modItem.LocalFullPath)
                        End If
                    Catch ex As Exception
                        modItem.InternalDetails.Add("Diferencia detectada en peso o fecha de modificación.")
                    End Try
                End Sub)
        End If

        ' Ordenar: Archivos que requieren acción primero, luego por ruta alfabética
        resultados.Sort(Function(a, b)
                            Dim rankA = If(a.Status = FileSyncStatusEnum.Identical, 1, 0)
                            Dim rankB = If(b.Status = FileSyncStatusEnum.Identical, 1, 0)
                            If rankA <> rankB Then Return rankA.CompareTo(rankB)
                            Return a.RelativePath.CompareTo(b.RelativePath)
                        End Function)

        progreso?.Report($"Comprobación finalizada. {resultados.Count} archivos procesados.")
        Return resultados
    End Function

    ''' <summary>
    ''' Escanea una carpeta completa utilizando DirectoryInfo.EnumerateFiles.
    ''' Obtiene Size y LastWriteTime en paquetes masivos desde el SMB Cache de Windows.
    ''' </summary>
    Private Function EscanearDirectorioRapido(directorioRaiz As String) As Dictionary(Of String, FileQuickMeta)
        Dim dicc As New Dictionary(Of String, FileQuickMeta)(StringComparer.OrdinalIgnoreCase)
        If Not Directory.Exists(directorioRaiz) Then Return dicc

        Try
            Dim dirInfo As New DirectoryInfo(directorioRaiz)
            Dim opt As New EnumerationOptions With {
                .RecurseSubdirectories = True,
                .IgnoreInaccessible = True,
                .ReturnSpecialDirectories = False
            }

            For Each fi In dirInfo.EnumerateFiles("*", opt)
                If EsArchivoIgnorado(fi.FullName) Then Continue For

                Dim relPath = GetRelativePath(directorioRaiz, fi.FullName)
                Dim meta As New FileQuickMeta With {
                    .RelativePath = relPath,
                    .FullPath = fi.FullName,
                    .Size = fi.Length,
                    .LastWriteUtc = fi.LastWriteTimeUtc
                }

                dicc(relPath) = meta
            Next
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error al escanear directorio {directorioRaiz}: {ex.Message}")
        End Try

        Return dicc
    End Function

    ''' <summary>
    ''' Copia de alta velocidad utilizando FileStream con búfer de 1 MB y FileOptions.SequentialScan
    ''' para saturar enlaces de red 5GbE / 10GbE.
    ''' </summary>
    Public Shared Sub CopiarArchivoStreamAltaVelocidad(origen As String, destino As String)
        Dim dirDestino = Path.GetDirectoryName(destino)
        If Not Directory.Exists(dirDestino) Then Directory.CreateDirectory(dirDestino)

        Const bufferSize As Integer = 1024 * 1024 ' 1 Megabyte por bloque

        Using sourceStream As New FileStream(origen, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, bufferSize, System.IO.FileOptions.SequentialScan)
            Using destStream As New FileStream(destino, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize, System.IO.FileOptions.WriteThrough)
                sourceStream.CopyTo(destStream, bufferSize)
            End Using
        End Using

        ' Preservar fecha de modificación original
        Try
            File.SetLastWriteTimeUtc(destino, File.GetLastWriteTimeUtc(origen))
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Ejecuta PULL en PARALELO: descarga los archivos más recientes del servidor a la carpeta local
    ''' utilizando transferencias concurrentes y búferes ampliados.
    ''' </summary>
    Public Function EjecutarPull(items As List(Of FileComparisonItem), localRoot As String, serverRoot As String, Optional progreso As IProgress(Of String) = Nothing) As Integer
        Dim itemsParaBajar = items.FindAll(Function(x) x.Status = FileSyncStatusEnum.ServerNewer OrElse x.Status = FileSyncStatusEnum.OnlyOnServer)
        If itemsParaBajar.Count = 0 Then Return 0

        Dim actualizados As Integer = 0
        Dim logCambios As New StringBuilder()

        logCambios.AppendLine("OPERACIÓN: PULL DE ALTA VELOCIDAD (Servidor -> Local)")
        logCambios.AppendLine($"Fecha/Hora: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
        logCambios.AppendLine($"Usuario local: {SysEnv.UserName} ({SysEnv.MachineName})")
        logCambios.AppendLine($"Total a sincronizar: {itemsParaBajar.Count} archivos")
        logCambios.AppendLine("--------------------------------------------------------------------------------")

        Dim maxHilos = Math.Min(6, Math.Max(2, SysEnv.ProcessorCount))
        Dim parOpts As New ParallelOptions With {.MaxDegreeOfParallelism = maxHilos}

        Dim swTotal = Stopwatch.StartNew()
        Dim bytesTransferidosTotal As Long = 0

        Parallel.ForEach(itemsParaBajar, parOpts,
            Sub(item As FileComparisonItem)
                Try
                    CopiarArchivoStreamAltaVelocidad(item.ServerFullPath, item.LocalFullPath)
                    Threading.Interlocked.Increment(actualizados)
                    Threading.Interlocked.Add(bytesTransferidosTotal, item.ServerSize)

                    SyncLock logCambios
                        logCambios.AppendLine($"• ACTUALIZADO: {item.RelativePath}")
                        logCambios.AppendLine($"  - Servidor: {item.ServerModified:yyyy-MM-dd HH:mm:ss} | {FormatBytes(item.ServerSize)}")
                        logCambios.AppendLine($"  - Local anterior: {If(item.LocalModified > DateTime.MinValue, item.LocalModified.ToString("yyyy-MM-dd HH:mm:ss"), "No existía")}")
                        If item.InternalDetails.Count > 0 Then
                            For Each det In item.InternalDetails
                                logCambios.AppendLine($"    * {det}")
                            Next
                        End If
                        logCambios.AppendLine()
                    End SyncLock

                    progreso?.Report($"Descargado [{actualizados}/{itemsParaBajar.Count}]: {item.RelativePath}")
                Catch ex As Exception
                    SyncLock logCambios
                        logCambios.AppendLine($"• ERROR al actualizar {item.RelativePath}: {ex.Message}")
                    End SyncLock
                End Try
            End Sub)

        swTotal.Stop()
        Dim segs = Math.Max(0.01, swTotal.Elapsed.TotalSeconds)
        Dim mbTransferidos = bytesTransferidosTotal / (1024.0 * 1024.0)
        Dim velocidadMbps = (mbTransferidos * 8.0) / segs

        logCambios.AppendLine("--------------------------------------------------------------------------------")
        logCambios.AppendLine($"Rendimiento: {FormatBytes(bytesTransferidosTotal)} transferidos en {segs:F2} seg ({velocidadMbps:F1} Mbps promedio)")

        If actualizados > 0 Then
            GuardarRegistroLog(serverRoot, localRoot, "PULL", logCambios.ToString())
        End If

        Return actualizados
    End Function

    ''' <summary>
    ''' Ejecuta PUSH en PARALELO: sube los archivos más recientes de la carpeta local al servidor
    ''' utilizando transferencias concurrentes y búferes ampliados.
    ''' </summary>
    Public Function EjecutarPush(items As List(Of FileComparisonItem), localRoot As String, serverRoot As String, Optional progreso As IProgress(Of String) = Nothing) As Integer
        Dim itemsParaSubir = items.FindAll(Function(x) x.Status = FileSyncStatusEnum.LocalNewer OrElse x.Status = FileSyncStatusEnum.OnlyOnLocal)
        If itemsParaSubir.Count = 0 Then Return 0

        Dim subidos As Integer = 0
        Dim logCambios As New StringBuilder()

        logCambios.AppendLine("OPERACIÓN: PUSH DE ALTA VELOCIDAD (Local -> Servidor)")
        logCambios.AppendLine($"Fecha/Hora: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
        logCambios.AppendLine($"Usuario autor: {SysEnv.UserName} ({SysEnv.MachineName})")
        logCambios.AppendLine($"Total a subir: {itemsParaSubir.Count} archivos")
        logCambios.AppendLine("--------------------------------------------------------------------------------")

        Dim maxHilos = Math.Min(6, Math.Max(2, SysEnv.ProcessorCount))
        Dim parOpts As New ParallelOptions With {.MaxDegreeOfParallelism = maxHilos}

        Dim swTotal = Stopwatch.StartNew()
        Dim bytesTransferidosTotal As Long = 0

        Parallel.ForEach(itemsParaSubir, parOpts,
            Sub(item As FileComparisonItem)
                Try
                    CopiarArchivoStreamAltaVelocidad(item.LocalFullPath, item.ServerFullPath)
                    Threading.Interlocked.Increment(subidos)
                    Threading.Interlocked.Add(bytesTransferidosTotal, item.LocalSize)

                    SyncLock logCambios
                        logCambios.AppendLine($"• SUBIDO AL SERVIDOR: {item.RelativePath}")
                        logCambios.AppendLine($"  - Fecha Local: {item.LocalModified:yyyy-MM-dd HH:mm:ss} | {FormatBytes(item.LocalSize)} | Por: {SysEnv.UserName}")
                        logCambios.AppendLine($"  - Servidor anterior: {If(item.ServerModified > DateTime.MinValue, item.ServerModified.ToString("yyyy-MM-dd HH:mm:ss"), "No existía en servidor")}")
                        If item.InternalDetails.Count > 0 Then
                            For Each det In item.InternalDetails
                                logCambios.AppendLine($"    * {det}")
                            Next
                        End If
                        logCambios.AppendLine()
                    End SyncLock

                    progreso?.Report($"Subido [{subidos}/{itemsParaSubir.Count}]: {item.RelativePath}")
                Catch ex As Exception
                    SyncLock logCambios
                        logCambios.AppendLine($"• ERROR al subir {item.RelativePath}: {ex.Message}")
                    End SyncLock
                End Try
            End Sub)

        swTotal.Stop()
        Dim segs = Math.Max(0.01, swTotal.Elapsed.TotalSeconds)
        Dim mbTransferidos = bytesTransferidosTotal / (1024.0 * 1024.0)
        Dim velocidadMbps = (mbTransferidos * 8.0) / segs

        logCambios.AppendLine("--------------------------------------------------------------------------------")
        logCambios.AppendLine($"Rendimiento: {FormatBytes(bytesTransferidosTotal)} transferidos en {segs:F2} seg ({velocidadMbps:F1} Mbps promedio)")

        If subidos > 0 Then
            GuardarRegistroLog(serverRoot, localRoot, "PUSH", logCambios.ToString())
        End If

        Return subidos
    End Function

    ''' <summary>
    ''' Inspecciona en profundidad el contenido interno de dos archivos (anterior y nuevo)
    ''' según su tipo (.adsklib / .asdklib para materiales, .vb/.txt para iLogic, .xml para estilos).
    ''' </summary>
    Public Function InspeccionarDiferenciasInternas(archivoAnterior As String, archivoNuevo As String) As List(Of String)
        Dim detalles As New List(Of String)

        If Not File.Exists(archivoAnterior) OrElse Not File.Exists(archivoNuevo) Then
            detalles.Add("No se puede comparar el contenido porque una de las versiones no está disponible.")
            Return detalles
        End If

        Dim extension = Path.GetExtension(archivoNuevo).ToLowerInvariant()

        Try
            Select Case extension
                Case ".adsklib", ".asdklib"
                    ' Inspección de bibliotecas de materiales de Autodesk
                    detalles.AddRange(CompararMaterialesAdsklib(archivoAnterior, archivoNuevo))

                Case ".vb", ".txt", ".iLogicVb"
                    ' Inspección de código de reglas iLogic
                    detalles.AddRange(CompararCodigoReglas(archivoAnterior, archivoNuevo))

                Case ".xml", ".styxml"
                    ' Inspección de estilos y plantillas XML
                    detalles.AddRange(CompararArchivosXml(archivoAnterior, archivoNuevo))

                Case ".xlsx"
                    Dim fiAnt As New FileInfo(archivoAnterior)
                    Dim fiNue As New FileInfo(archivoNuevo)
                    Dim delta = fiNue.Length - fiAnt.Length
                    detalles.Add($"Planilla Excel de requerimientos actualizada. Diferencia de peso: {If(delta >= 0, "+" & FormatBytes(delta), FormatBytes(delta))}.")

                Case Else
                    Dim fiAnt As New FileInfo(archivoAnterior)
                    Dim fiNue As New FileInfo(archivoNuevo)
                    Dim delta = fiNue.Length - fiAnt.Length
                    detalles.Add($"Variación de tamaño: {If(delta >= 0, "+" & FormatBytes(delta), FormatBytes(delta))}.")
            End Select
        Catch ex As Exception
            detalles.Add($"No se pudo inspeccionar contenido interno: {ex.Message}")
        End Try

        Return detalles
    End Function

    ''' <summary>
    ''' Estructura que separa rigurosamente los Materiales físicos de los Aspectos / Apariencias
    ''' visuales dentro de una biblioteca .adsklib de Autodesk Inventor.
    ''' </summary>
    Public Class BibliotecaAssetsInfo
        Public Property Materiales As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Public Property Aspectos As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Public Property MetodoLectura As String = ""
    End Class

    ''' <summary>
    ''' Extrae y compara la lista de materiales y aspectos dentro de archivos .adsklib / .asdklib.
    ''' Procesa de manera simétrica y categorizada ambos archivos (local y servidor) garantizando
    ''' la detección exacta de altas, bajas (eliminaciones) de Materiales (ej: SAE 1045) y de Aspectos.
    ''' </summary>
    Private Function CompararMaterialesAdsklib(rutaAnterior As String, rutaNueva As String) As List(Of String)
        Dim diff As New List(Of String)

        Dim infoAnt = ObtenerContenidoBiblioteca(rutaAnterior)
        Dim infoNue = ObtenerContenidoBiblioteca(rutaNueva)

        ' Homogeneización inteligente cruzada:
        ' Si un nombre fue identificado inequívocamente como Aspecto en una de las fuentes (ej. Inventor API),
        ' clasificarlo como Aspecto en ambas para evitar reclasificaciones erróneas entre versiones.
        Dim aspectosConocidos As New HashSet(Of String)(infoAnt.Aspectos, StringComparer.OrdinalIgnoreCase)
        For Each asp In infoNue.Aspectos : aspectosConocidos.Add(asp) : Next

        For Each asp In aspectosConocidos
            If infoAnt.Materiales.Contains(asp) Then
                infoAnt.Materiales.Remove(asp)
                infoAnt.Aspectos.Add(asp)
            End If
            If infoNue.Materiales.Contains(asp) Then
                infoNue.Materiales.Remove(asp)
                infoNue.Aspectos.Add(asp)
            End If
        Next

        Dim nombreLib = Path.GetFileName(rutaNueva)
        diff.Add($"Inspección de Biblioteca de Materiales: ""{nombreLib}""")

        If infoAnt.MetodoLectura = infoNue.MetodoLectura Then
            diff.Add($"• Método de lectura: {infoNue.MetodoLectura}")
        Else
            diff.Add($"• Método de lectura: {infoNue.MetodoLectura} (actual) / {infoAnt.MetodoLectura} (anterior)")
        End If

        ' --- 1. SECCIÓN MATERIALES FÍSICOS ---
        Dim matsAgregados As New List(Of String)
        For Each m In infoNue.Materiales
            If Not infoAnt.Materiales.Contains(m) Then matsAgregados.Add(m)
        Next

        Dim matsEliminados As New List(Of String)
        For Each m In infoAnt.Materiales
            If Not infoNue.Materiales.Contains(m) Then matsEliminados.Add(m)
        Next

        matsAgregados.Sort(StringComparer.OrdinalIgnoreCase)
        matsEliminados.Sort(StringComparer.OrdinalIgnoreCase)

        diff.Add("")
        diff.Add("================================================================================")
        diff.Add("📦 SECCIÓN 1: MATERIALES FÍSICOS (MaterialAssets)")
        diff.Add($"• Cantidad de materiales: {infoAnt.Materiales.Count} (anterior) -> {infoNue.Materiales.Count} (nuevo)")

        If matsAgregados.Count > 0 Then
            diff.Add($"• Materiales nuevos agregados ({matsAgregados.Count}):")
            For Each m In matsAgregados
                diff.Add($"    [+] Material: ""{m}""")
            Next
        End If

        If matsEliminados.Count > 0 Then
            diff.Add($"• Materiales eliminados ({matsEliminados.Count}):")
            For Each m In matsEliminados
                diff.Add($"    [-] Material eliminado: ""{m}""")
            Next
        End If

        If matsAgregados.Count = 0 AndAlso matsEliminados.Count = 0 Then
            diff.Add("• Sin altas ni bajas de materiales (listado 100% coincidente).")
        End If

        ' --- 2. SECCIÓN ASPECTOS Y APARIENCIAS ---
        Dim aspAgregados As New List(Of String)
        For Each a In infoNue.Aspectos
            If Not infoAnt.Aspectos.Contains(a) Then aspAgregados.Add(a)
        Next

        Dim aspEliminados As New List(Of String)
        For Each a In infoAnt.Aspectos
            If Not infoNue.Aspectos.Contains(a) Then aspEliminados.Add(a)
        Next

        aspAgregados.Sort(StringComparer.OrdinalIgnoreCase)
        aspEliminados.Sort(StringComparer.OrdinalIgnoreCase)

        diff.Add("")
        diff.Add("================================================================================")
        diff.Add("🎨 SECCIÓN 2: ASPECTOS Y APARIENCIAS (AppearanceAssets)")
        diff.Add($"• Cantidad de apariencias: {infoAnt.Aspectos.Count} (anterior) -> {infoNue.Aspectos.Count} (nuevo)")

        If aspAgregados.Count > 0 Then
            diff.Add($"• Aspectos / Apariencias nuevas agregadas ({aspAgregados.Count}):")
            For Each a In aspAgregados
                diff.Add($"    [+] Aspecto: ""{a}""")
            Next
        End If

        If aspEliminados.Count > 0 Then
            diff.Add($"• Aspectos / Apariencias eliminadas ({aspEliminados.Count}):")
            For Each a In aspEliminados
                diff.Add($"    [-] Aspecto eliminado: ""{a}""")
            Next
        End If

        If aspAgregados.Count = 0 AndAlso aspEliminados.Count = 0 Then
            diff.Add("• Sin altas ni bajas de aspectos (listado 100% coincidente).")
        End If

        If matsAgregados.Count = 0 AndAlso matsEliminados.Count = 0 AndAlso
           aspAgregados.Count = 0 AndAlso aspEliminados.Count = 0 Then
            Dim fiAnt As New FileInfo(rutaAnterior)
            Dim fiNue As New FileInfo(rutaNueva)
            Dim delta = fiNue.Length - fiAnt.Length
            diff.Add("")
            diff.Add($"• Modificación interna: Se actualizaron propiedades físicas (densidad, límite elástico) o mapas gráficos ({If(delta >= 0, "+" & FormatBytes(delta), FormatBytes(delta))}).")
        End If

        Return diff
    End Function

    ''' <summary>
    ''' Obtiene el conjunto de materiales y aspectos de una biblioteca .adsklib/.asdklib.
    ''' Prioriza la API oficial de Inventor si la biblioteca está cargada en la sesión,
    ''' y recurre a un extractor estructural directo del paquete (.adsklib) que funciona en red UNC sin COM.
    ''' </summary>
    Public Function ObtenerContenidoBiblioteca(rutaAdsklib As String) As BibliotecaAssetsInfo
        ' 1. Intentar consultar vía Autodesk Inventor API si está disponible en la sesión
        If _invApp IsNot Nothing Then
            Try
                Dim infoInv = ExtraerAssetsViaInventor(rutaAdsklib)
                If infoInv IsNot Nothing AndAlso (infoInv.Materiales.Count > 0 OrElse infoInv.Aspectos.Count > 0) Then
                    Return infoInv
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Extracción vía Inventor API no disponible para este archivo: " & ex.Message)
            End Try
        End If

        ' 2. Extractor estructural directo del paquete ZIP / XML (inmune a restricciones COM de red)
        Return ExtraerAssetsDesdePaquete(rutaAdsklib)
    End Function

    ''' <summary>
    ''' Consulta las colecciones MaterialAssets y AppearanceAssets de Autodesk Inventor.
    ''' </summary>
    Private Function ExtraerAssetsViaInventor(rutaAdsklib As String) As BibliotecaAssetsInfo
        Dim info As New BibliotecaAssetsInfo()
        Dim inv = ObtenerInventorApp()
        If inv Is Nothing Then Return info

        Try
            Dim assetLib As AssetLibrary = Nothing
            Dim nombreBase = Path.GetFileNameWithoutExtension(rutaAdsklib)
            Dim fullPathNorm = Path.GetFullPath(rutaAdsklib)

            ' 1. Buscar en librerías abiertas de Inventor por DisplayName, Name o Ruta
            Try
                For Each al As AssetLibrary In inv.AssetLibraries
                    Try
                        Dim match As Boolean = False
                        Try
                            If String.Equals(al.DisplayName, nombreBase, StringComparison.OrdinalIgnoreCase) OrElse
                               String.Equals(Path.GetFileNameWithoutExtension(al.DisplayName), nombreBase, StringComparison.OrdinalIgnoreCase) Then
                                match = True
                            End If
                        Catch
                        End Try

                        If Not match Then
                            Try
                                If String.Equals(al.Name, nombreBase, StringComparison.OrdinalIgnoreCase) OrElse
                                   String.Equals(Path.GetFileNameWithoutExtension(al.Name), nombreBase, StringComparison.OrdinalIgnoreCase) Then
                                    match = True
                                End If
                            Catch
                            End Try
                        End If

                        If Not match Then
                            Try
                                If String.Equals(Path.GetFileNameWithoutExtension(al.FullFileName), nombreBase, StringComparison.OrdinalIgnoreCase) Then
                                    match = True
                                End If
                            Catch
                            End Try
                        End If

                        If Not match Then
                            Try
                                If String.Equals(Path.GetFullPath(al.FullFileName), fullPathNorm, StringComparison.OrdinalIgnoreCase) Then
                                    match = True
                                End If
                            Catch
                            End Try
                        End If

                        If match Then
                            assetLib = al
                            Exit For
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            If assetLib Is Nothing Then
                Try
                    If File.Exists(rutaAdsklib) Then
                        assetLib = inv.AssetLibraries.Open(rutaAdsklib)
                    End If
                Catch
                End Try
            End If

            If assetLib IsNot Nothing Then
                info.MetodoLectura = "Autodesk Inventor AssetLibraries API"

                ' Materiales Físicos (kAssetTypeMaterial)
                Try
                    If assetLib.MaterialAssets IsNot Nothing Then
                        For Each asset As Asset In assetLib.MaterialAssets
                            Try
                                Dim nombre = asset.DisplayName
                                If String.IsNullOrWhiteSpace(nombre) Then nombre = asset.Name
                                If Not String.IsNullOrWhiteSpace(nombre) AndAlso nombre.IndexOf(ChrW(&HFFFD)) < 0 AndAlso Not Regex.IsMatch(nombre, "^\{?[0-9a-fA-F]{4,8}(-[0-9a-fA-F]{4,8}){2,5}\}?$") Then
                                    info.Materiales.Add(nombre.Trim())
                                End If
                            Catch
                            End Try
                        Next
                    End If
                Catch
                End Try

                ' Aspectos y Apariencias (kAssetTypeAppearance)
                Try
                    If assetLib.AppearanceAssets IsNot Nothing Then
                        For Each appAsset As Asset In assetLib.AppearanceAssets
                            Try
                                Dim nombre = appAsset.DisplayName
                                If String.IsNullOrWhiteSpace(nombre) Then nombre = appAsset.Name
                                If Not String.IsNullOrWhiteSpace(nombre) AndAlso nombre.IndexOf(ChrW(&HFFFD)) < 0 AndAlso Not Regex.IsMatch(nombre, "^\{?[0-9a-fA-F]{4,8}(-[0-9a-fA-F]{4,8}){2,5}\}?$") Then
                                    info.Aspectos.Add(nombre.Trim())
                                End If
                            Catch
                            End Try
                        Next
                    End If
                Catch
                End Try
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error en ExtraerAssetsViaInventor: " & ex.Message)
        End Try

        Return info
    End Function

    ''' <summary>
    ''' Ejecutado en el hilo STA / UI donde la COM API de Inventor está activa sin bloqueos RPC.
    ''' Consulta las colecciones nativas MaterialAssets y AppearanceAssets de Autodesk Inventor
    ''' y las compara contra el archivo de servidor para reportar con precisión absoluta
    ''' las cantidades exactas (ej: 9 -> 10 materiales) y nombres reales de altas y bajas.
    ''' </summary>
    Public Sub RefinarInspeccionConInventor(item As FileComparisonItem)
        If item Is Nothing Then Exit Sub
        Dim ext = Path.GetExtension(item.RelativePath).ToLowerInvariant()
        If ext <> ".adsklib" AndAlso ext <> ".asdklib" Then Exit Sub

        Try
            Dim inv = ObtenerInventorApp()
            Dim nombreLib = Path.GetFileNameWithoutExtension(item.RelativePath)
            Dim assetLib As AssetLibrary = Nothing

            ' 1. Localizar la librería en la sesión activa de Inventor si está disponible
            If inv IsNot Nothing Then
                Try
                    For Each al As AssetLibrary In inv.AssetLibraries
                        Try
                            Dim match As Boolean = False
                            Try
                                If String.Equals(al.DisplayName, nombreLib, StringComparison.OrdinalIgnoreCase) OrElse
                                   String.Equals(Path.GetFileNameWithoutExtension(al.DisplayName), nombreLib, StringComparison.OrdinalIgnoreCase) Then
                                    match = True
                                End If
                            Catch
                            End Try

                            If Not match Then
                                Try
                                    If String.Equals(al.Name, nombreLib, StringComparison.OrdinalIgnoreCase) OrElse
                                       String.Equals(Path.GetFileNameWithoutExtension(al.Name), nombreLib, StringComparison.OrdinalIgnoreCase) Then
                                        match = True
                                    End If
                                Catch
                                End Try
                            End If

                            If Not match Then
                                Try
                                    If String.Equals(Path.GetFileNameWithoutExtension(al.FullFileName), nombreLib, StringComparison.OrdinalIgnoreCase) Then
                                        match = True
                                    End If
                                Catch
                                End Try
                            End If

                            If Not match AndAlso File.Exists(item.LocalFullPath) Then
                                Try
                                    If String.Equals(Path.GetFullPath(al.FullFileName), Path.GetFullPath(item.LocalFullPath), StringComparison.OrdinalIgnoreCase) Then
                                        match = True
                                    End If
                                Catch
                                End Try
                            End If

                            If match Then
                                assetLib = al
                                Exit For
                            End If
                        Catch
                        End Try
                    Next
                Catch
                End Try

                If assetLib Is Nothing Then
                    Try
                        If File.Exists(item.LocalFullPath) Then
                            assetLib = inv.AssetLibraries.Open(item.LocalFullPath)
                        End If
                    Catch
                    End Try
                End If
            End If

            ' 2. Extraer listas oficiales desde la API de Inventor (Local) o desde paquete
            Dim matsLocal As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim aspsLocal As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim metodoLectura As String = "Inspección estructural de biblioteca (.adsklib)"

            If assetLib IsNot Nothing Then
                metodoLectura = "Autodesk Inventor AssetLibraries API"

                Try
                    If assetLib.MaterialAssets IsNot Nothing Then
                        For Each asset As Asset In assetLib.MaterialAssets
                            Try
                                Dim nom = asset.DisplayName
                                If String.IsNullOrWhiteSpace(nom) Then nom = asset.Name
                                If Not String.IsNullOrWhiteSpace(nom) AndAlso nom.IndexOf(ChrW(&HFFFD)) < 0 AndAlso Not Regex.IsMatch(nom, "^\{?[0-9a-fA-F]{4,8}(-[0-9a-fA-F]{4,8}){2,5}\}?$") Then
                                    matsLocal.Add(nom.Trim())
                                End If
                            Catch
                            End Try
                        Next
                    End If
                Catch
                End Try

                Try
                    If assetLib.AppearanceAssets IsNot Nothing Then
                        For Each appAsset As Asset In assetLib.AppearanceAssets
                            Try
                                Dim nom = appAsset.DisplayName
                                If String.IsNullOrWhiteSpace(nom) Then nom = appAsset.Name
                                If Not String.IsNullOrWhiteSpace(nom) AndAlso nom.IndexOf(ChrW(&HFFFD)) < 0 AndAlso Not Regex.IsMatch(nom, "^\{?[0-9a-fA-F]{4,8}(-[0-9a-fA-F]{4,8}){2,5}\}?$") Then
                                    aspsLocal.Add(nom.Trim())
                                End If
                            Catch
                            End Try
                        Next
                    End If
                Catch
                End Try
            ElseIf File.Exists(item.LocalFullPath) Then
                ' Fallback: extraer directamente del paquete local
                Dim infoLocalPkg = ExtraerAssetsDesdePaquete(item.LocalFullPath)
                For Each m In infoLocalPkg.Materiales : matsLocal.Add(m) : Next
                For Each a In infoLocalPkg.Aspectos : aspsLocal.Add(a) : Next
            End If

            If matsLocal.Count = 0 AndAlso aspsLocal.Count = 0 Then Exit Sub

            ' 3. Determinar qué contiene la versión del servidor
            Dim matsServer As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim aspsServer As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If File.Exists(item.ServerFullPath) Then
                ' Verificar presencia exacta en el paquete del servidor
                VerificarPresenciaEnPaquete(item.ServerFullPath, matsLocal, aspsLocal, matsServer, aspsServer)

                ' También verificar si había materiales/aspectos existentes en el servidor que ya no están en local
                Try
                    Dim infoServerPkg = ExtraerAssetsDesdePaquete(item.ServerFullPath)
                    For Each m In infoServerPkg.Materiales
                        If Not matsLocal.Contains(m) AndAlso EsNombreAssetValido(m) Then matsServer.Add(m)
                    Next
                    For Each a In infoServerPkg.Aspectos
                        If Not aspsLocal.Contains(a) AndAlso EsNombreAssetValido(a) Then aspsServer.Add(a)
                    Next
                Catch
                End Try
            End If

            ' 4. Construir el reporte comparativo exacto según si es PUSH (servidor anterior, local nuevo) o PULL
            Dim matsAnt As HashSet(Of String)
            Dim matsNue As HashSet(Of String)
            Dim aspsAnt As HashSet(Of String)
            Dim aspsNue As HashSet(Of String)

            If item.Status = FileSyncStatusEnum.LocalNewer Then
                ' PUSH: Servidor es la versión anterior, Local es la nueva
                matsAnt = matsServer
                matsNue = matsLocal
                aspsAnt = aspsServer
                aspsNue = aspsLocal
            Else
                ' PULL: Local es la versión anterior, Servidor es la nueva
                matsAnt = matsLocal
                matsNue = matsServer
                aspsAnt = aspsLocal
                aspsNue = aspsServer
            End If

            Dim nuevosDetalles As New List(Of String)()
            nuevosDetalles.Add($"Inspección de Biblioteca de Materiales: ""{Path.GetFileName(item.RelativePath)}""")
            nuevosDetalles.Add("• Método de lectura: " & metodoLectura)

            ' --- SECCIÓN 1: MATERIALES FÍSICOS ---
            Dim matsAgregados As New List(Of String)()
            For Each m In matsNue
                If Not matsAnt.Contains(m) Then matsAgregados.Add(m)
            Next

            Dim matsEliminados As New List(Of String)()
            For Each m In matsAnt
                If Not matsNue.Contains(m) Then matsEliminados.Add(m)
            Next

            matsAgregados.Sort(StringComparer.OrdinalIgnoreCase)
            matsEliminados.Sort(StringComparer.OrdinalIgnoreCase)

            nuevosDetalles.Add("")
            nuevosDetalles.Add("================================================================================")
            nuevosDetalles.Add("📦 SECCIÓN 1: MATERIALES FÍSICOS (MaterialAssets)")
            nuevosDetalles.Add($"• Cantidad de materiales: {matsAnt.Count} (anterior) -> {matsNue.Count} (nuevo)")

            If matsAgregados.Count > 0 Then
                nuevosDetalles.Add($"• Materiales nuevos agregados ({matsAgregados.Count}):")
                For Each m In matsAgregados
                    nuevosDetalles.Add($"    [+] Material: ""{m}""")
                Next
            End If

            If matsEliminados.Count > 0 Then
                nuevosDetalles.Add($"• Materiales eliminados ({matsEliminados.Count}):")
                For Each m In matsEliminados
                    nuevosDetalles.Add($"    [-] Material eliminado: ""{m}""")
                Next
            End If

            If matsAgregados.Count = 0 AndAlso matsEliminados.Count = 0 Then
                nuevosDetalles.Add("• Sin altas ni bajas de materiales (listado 100% coincidente).")
            End If

            ' --- SECCIÓN 2: ASPECTOS Y APARIENCIAS ---
            Dim aspsAgregados As New List(Of String)()
            For Each a In aspsNue
                If Not aspsAnt.Contains(a) Then aspsAgregados.Add(a)
            Next

            Dim aspsEliminados As New List(Of String)()
            For Each a In aspsAnt
                If Not aspsNue.Contains(a) Then aspsEliminados.Add(a)
            Next

            aspsAgregados.Sort(StringComparer.OrdinalIgnoreCase)
            aspsEliminados.Sort(StringComparer.OrdinalIgnoreCase)

            nuevosDetalles.Add("")
            nuevosDetalles.Add("================================================================================")
            nuevosDetalles.Add("🎨 SECCIÓN 2: ASPECTOS Y APARIENCIAS (AppearanceAssets)")
            nuevosDetalles.Add($"• Cantidad de apariencias: {aspsAnt.Count} (anterior) -> {aspsNue.Count} (nuevo)")

            If aspsAgregados.Count > 0 Then
                nuevosDetalles.Add($"• Aspectos / Apariencias nuevas agregadas ({aspsAgregados.Count}):")
                For Each a In aspsAgregados
                    nuevosDetalles.Add($"    [+] Aspecto: ""{a}""")
                Next
            End If

            If aspsEliminados.Count > 0 Then
                nuevosDetalles.Add($"• Aspectos / Apariencias eliminadas ({aspsEliminados.Count}):")
                For Each a In aspsEliminados
                    nuevosDetalles.Add($"    [-] Aspecto eliminado: ""{a}""")
                Next
            End If

            If aspsAgregados.Count = 0 AndAlso aspsEliminados.Count = 0 Then
                nuevosDetalles.Add("• Sin altas ni bajas de aspectos (listado 100% coincidente).")
            End If

            If matsAgregados.Count = 0 AndAlso matsEliminados.Count = 0 AndAlso
               aspsAgregados.Count = 0 AndAlso aspsEliminados.Count = 0 Then
                Dim delta = item.LocalSize - item.ServerSize
                nuevosDetalles.Add("")
                nuevosDetalles.Add($"• Modificación interna: Se actualizaron propiedades físicas (densidad, límite elástico) o mapas gráficos ({If(delta >= 0, "+" & FormatBytes(delta), FormatBytes(delta))}).")
            End If

            item.InternalDetails = nuevosDetalles

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error en RefinarInspeccionConInventor: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Verifica qué materiales y aspectos conocidos están presentes dentro de un archivo .adsklib remoto/local
    ''' buscando sus firmas de texto (UTF-8 y UTF-16) dentro del paquete ZIP.
    ''' </summary>
    Private Sub VerificarPresenciaEnPaquete(rutaAdsklib As String,
                                            matsCandidatos As HashSet(Of String),
                                            aspsCandidatos As HashSet(Of String),
                                            matsPresentes As HashSet(Of String),
                                            aspsPresentes As HashSet(Of String))
        If Not File.Exists(rutaAdsklib) Then Exit Sub

        Dim patronesMats As New Dictionary(Of String, List(Of Byte()))(StringComparer.OrdinalIgnoreCase)
        For Each m In matsCandidatos
            Dim listaBytes As New List(Of Byte()) From {
                Encoding.UTF8.GetBytes(m),
                Encoding.Unicode.GetBytes(m)
            }
            patronesMats(m) = listaBytes
        Next

        Dim patronesAsps As New Dictionary(Of String, List(Of Byte()))(StringComparer.OrdinalIgnoreCase)
        For Each a In aspsCandidatos
            Dim listaBytes As New List(Of Byte()) From {
                Encoding.UTF8.GetBytes(a),
                Encoding.Unicode.GetBytes(a)
            }
            patronesAsps(a) = listaBytes
        Next

        Try
            Using fs As New FileStream(rutaAdsklib, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using zip As New ZipArchive(fs, ZipArchiveMode.Read)
                    For Each entry In zip.Entries
                        Dim ext = Path.GetExtension(entry.Name).ToLowerInvariant()
                        If ext = ".dds" OrElse ext = ".png" OrElse ext = ".jpg" OrElse ext = ".bmp" OrElse
                           ext = ".tif" OrElse ext = ".tga" OrElse ext = ".fx" OrElse ext = ".hlsl" OrElse
                           ext = ".ico" OrElse ext = ".jpeg" OrElse ext = ".cso" Then
                            Continue For
                        End If

                        Try
                            Using s = entry.Open()
                                Using ms As New MemoryStream()
                                    s.CopyTo(ms)
                                    Dim fileBytes = ms.ToArray()

                                    For Each kvp In patronesMats
                                        If Not matsPresentes.Contains(kvp.Key) Then
                                            For Each pat In kvp.Value
                                                If BuscarSecuenciaBytes(fileBytes, pat) Then
                                                    matsPresentes.Add(kvp.Key)
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    Next

                                    For Each kvp In patronesAsps
                                        If Not aspsPresentes.Contains(kvp.Key) Then
                                            For Each pat In kvp.Value
                                                If BuscarSecuenciaBytes(fileBytes, pat) Then
                                                    aspsPresentes.Add(kvp.Key)
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    Next
                                End Using
                            End Using
                        Catch
                        End Try

                        If matsPresentes.Count = matsCandidatos.Count AndAlso aspsPresentes.Count = aspsCandidatos.Count Then
                            Exit For
                        End If
                    Next
                End Using
            End Using
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error en VerificarPresenciaEnPaquete: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Búsqueda veloz de una secuencia de bytes dentro de un array.
    ''' </summary>
    Private Function BuscarSecuenciaBytes(origen As Byte(), patron As Byte()) As Boolean
        If origen Is Nothing OrElse patron Is Nothing OrElse patron.Length = 0 OrElse origen.Length < patron.Length Then
            Return False
        End If

        Dim primerByte = patron(0)
        Dim finBusqueda = origen.Length - patron.Length

        For i As Integer = 0 To finBusqueda
            If origen(i) = primerByte Then
                Dim coincide = True
                For j As Integer = 1 To patron.Length - 1
                    If origen(i + j) <> patron(j) Then
                        coincide = False
                        Exit For
                    End If
                Next
                If coincide Then Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' Extractor estructural profundo de bibliotecas .adsklib.
    ''' Inspecciona el paquete ZIP / XML directo (útil cuando la API COM de Inventor no está activa).
    ''' </summary>
    Private Function ExtraerAssetsDesdePaquete(rutaAdsklib As String) As BibliotecaAssetsInfo
        Dim info As New BibliotecaAssetsInfo() With {
            .MetodoLectura = "Inspección estructural de biblioteca (.adsklib)"
        }

        If Not File.Exists(rutaAdsklib) Then Return info

        Try
            Using fs As New FileStream(rutaAdsklib, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using zip As New ZipArchive(fs, ZipArchiveMode.Read)
                    For Each entry In zip.Entries
                        ' OMITIR únicamente imágenes, mapas de texturas y shaders binarios
                        Dim ext = Path.GetExtension(entry.Name).ToLowerInvariant()
                        If ext = ".dds" OrElse ext = ".png" OrElse ext = ".jpg" OrElse ext = ".bmp" OrElse
                           ext = ".tif" OrElse ext = ".tga" OrElse ext = ".fx" OrElse ext = ".hlsl" OrElse
                           ext = ".ico" OrElse ext = ".jpeg" OrElse ext = ".cso" Then
                            Continue For
                        End If

                        Dim isXml = ext = ".xml"

                        Try
                            Using s = entry.Open()
                                Using ms As New MemoryStream()
                                    s.CopyTo(ms)
                                    Dim bytes = ms.ToArray()

                                    If isXml Then
                                        Try
                                            Dim xmlText = Encoding.UTF8.GetString(bytes)
                                            ExtraerAssetsDesdeXml(xmlText, info)
                                        Catch
                                        End Try
                                    Else
                                        ExtraerAssetsDesdeBinario(bytes, info)
                                    End If
                                End Using
                            End Using
                        Catch
                        End Try
                    Next
                End Using
            End Using
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error en ExtraerAssetsDesdePaquete: " & ex.Message)
        End Try

        Return info
    End Function

    Private Sub ExtraerAssetsDesdeXml(xmlContent As String, info As BibliotecaAssetsInfo)
        If String.IsNullOrWhiteSpace(xmlContent) Then Return

        ' Buscar nodos <asset ...> o <item ...>
        Dim regexTags As New Regex("<(?:asset|assetDefinition|item)[^>]+>", RegexOptions.IgnoreCase)
        For Each matchTag As Match In regexTags.Matches(xmlContent)
            Dim tag = matchTag.Value
            Dim isAppearance = tag.IndexOf("appearance", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                               tag.IndexOf("prismopaque", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                               tag.IndexOf("prismmetal", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                               tag.IndexOf("generic", StringComparison.OrdinalIgnoreCase) >= 0

            Dim regexName As New Regex("(?:userLabel|displayName|assetLabel|name)=""([^""]+)""", RegexOptions.IgnoreCase)
            Dim mName = regexName.Match(tag)
            If mName.Success Then
                Dim nom = mName.Groups(1).Value.Trim()
                If EsNombreAssetValido(nom) Then
                    If isAppearance Then
                        info.Aspectos.Add(nom)
                    Else
                        info.Materiales.Add(nom)
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub ExtraerAssetsDesdeBinario(bytes As Byte(), info As BibliotecaAssetsInfo)
        If bytes Is Nothing OrElse bytes.Length < 16 Then Return

        Dim listaStrings = ExtraerStringsBinarios(bytes)
        If listaStrings.Count = 0 Then Return

        ' Clasificar según contexto en la tabla de contenidos
        Dim categoriaActualEsAspecto As Boolean = False

        For Each textoCadena As String In listaStrings
            Dim sLower = textoCadena.ToLowerInvariant()

            ' Detectar cambios de contexto de esquema
            If sLower.Contains("appearance") OrElse sLower.Contains("prismopaque") OrElse
               sLower.Contains("prismmetal") OrElse sLower.Contains("prismglazing") Then
                categoriaActualEsAspecto = True
            ElseIf sLower.Contains("material") OrElse sLower.Contains("physmat") Then
                categoriaActualEsAspecto = False
            End If

            If EsNombreAssetValido(textoCadena) Then
                Dim esAspectoEvidente = categoriaActualEsAspecto OrElse
                                        textoCadena.Contains(" - ") OrElse
                                        textoCadena.EndsWith(")") OrElse
                                        sLower.Contains("satinado") OrElse
                                        sLower.Contains("pulido") OrElse
                                        sLower.Contains("semipulido") OrElse
                                        sLower.Contains("cromo") OrElse
                                        sLower.Contains("gris hierro")

                If esAspectoEvidente Then
                    info.Aspectos.Add(textoCadena)
                Else
                    info.Materiales.Add(textoCadena)
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Extrae tanto cadenas UTF-8/ASCII como secuencias UTF-16 LE (Unicode) de un búfer binario de Autodesk Protein.
    ''' </summary>
    Private Function ExtraerStringsBinarios(bytes As Byte()) As List(Of String)
        Dim resultado As New List(Of String)()
        If bytes Is Nothing OrElse bytes.Length < 4 Then Return resultado

        ' 1. Cadenas UTF-8 / ASCII
        Dim bufAscii As New List(Of Byte)()
        For i As Integer = 0 To bytes.Length - 1
            Dim b = bytes(i)
            If (b >= 32 AndAlso b <= 126) OrElse (b >= 160 AndAlso b <= 255) Then
                bufAscii.Add(b)
            Else
                If bufAscii.Count >= 2 AndAlso bufAscii.Count <= 80 Then
                    Try
                        Dim s = Encoding.UTF8.GetString(bufAscii.ToArray()).Trim()
                        If Not String.IsNullOrEmpty(s) Then resultado.Add(s)
                    Catch
                    End Try
                End If
                bufAscii.Clear()
            End If
        Next
        If bufAscii.Count >= 2 AndAlso bufAscii.Count <= 80 Then
            Try
                Dim s = Encoding.UTF8.GetString(bufAscii.ToArray()).Trim()
                If Not String.IsNullOrEmpty(s) Then resultado.Add(s)
            Catch
            End Try
        End If

        ' 2. Cadenas UTF-16 Little Endian (Unicode 2 bytes por carácter)
        Dim bufUnicode As New List(Of Byte)()
        Dim j As Integer = 0
        While j < bytes.Length - 1
            Dim b1 = bytes(j)
            Dim b2 = bytes(j + 1)
            If b2 = 0 AndAlso ((b1 >= 32 AndAlso b1 <= 126) OrElse (b1 >= 160 AndAlso b1 <= 255)) Then
                bufUnicode.Add(b1)
                bufUnicode.Add(b2)
                j += 2
            Else
                If bufUnicode.Count >= 4 AndAlso bufUnicode.Count <= 160 Then
                    Try
                        Dim s = Encoding.Unicode.GetString(bufUnicode.ToArray()).Trim()
                        If Not String.IsNullOrEmpty(s) Then resultado.Add(s)
                    Catch
                    End Try
                End If
                bufUnicode.Clear()
                j += 1
            End If
        End While
        If bufUnicode.Count >= 4 AndAlso bufUnicode.Count <= 160 Then
            Try
                Dim s = Encoding.Unicode.GetString(bufUnicode.ToArray()).Trim()
                If Not String.IsNullOrEmpty(s) Then resultado.Add(s)
            Catch
            End Try
        End If

        Return resultado
    End Function

    ''' <summary>
    ''' Filtra nombres de materiales y aspectos reales, descartando estrictamente identificadores técnicos,
    ''' GUIDs de Protein, esquemas XML, tokens de shaders y secuencias corruptas.
    ''' </summary>
    Private Function EsNombreAssetValido(val As String) As Boolean
        If String.IsNullOrWhiteSpace(val) Then Return False
        val = val.Trim()
        If val.Length < 2 OrElse val.Length > 80 Then Return False

        ' Descartar caracteres corruptos o de reemplazo Unicode (\uFFFD) y nulos
        If val.IndexOf(ChrW(&HFFFD)) >= 0 OrElse val.IndexOf(vbNullChar) >= 0 Then Return False
        For Each ch As Char In val
            If Char.IsControl(ch) Then Return False
        Next

        ' Descartar GUIDs y hashes de Autodesk Protein (ej: "8f28bc60-41e1780e-1c5603bb-fe92204d")
        If Regex.IsMatch(val, "^\{?[0-9a-fA-F]{4,8}(-[0-9a-fA-F]{4,8}){2,5}\}?$") Then Return False

        ' Descartar tokens que sean secuencias hexadecimales largas sin espacios
        If Not val.Contains(" ") AndAlso Regex.IsMatch(val, "^[0-9a-fA-F\-]{16,}$") Then Return False

        Dim valLower = val.ToLowerInvariant()

        ' Descartar URLs, URNs y namespaces
        If valLower.StartsWith("http:") OrElse valLower.StartsWith("https:") OrElse
           valLower.StartsWith("urn:") OrElse valLower.StartsWith("xmlns") OrElse
           valLower.StartsWith("uuid:") OrElse valLower.Contains("schemas.autodesk") OrElse
           valLower.Contains("schemas.microsoft") OrElse valLower.Contains("w3.org") Then
            Return False
        End If

        ' Descartar extensiones internas
        If valLower.EndsWith(".xml") OrElse valLower.EndsWith(".bin") OrElse
           valLower.EndsWith(".inf") OrElse valLower.EndsWith(".png") OrElse
           valLower.EndsWith(".jpg") OrElse valLower.EndsWith(".dds") OrElse
           valLower.EndsWith(".bmp") OrElse valLower.EndsWith(".dat") Then
            Return False
        End If

        ' Descartar identificadores de esquemas de Autodesk Protein / Prism
        If valLower.Contains("schema") OrElse valLower.StartsWith("unifiedbitmap") OrElse
           valLower.StartsWith("generic_") OrElse valLower.StartsWith("opaque_") OrElse
           valLower.StartsWith("metal_") OrElse valLower.StartsWith("structural_") OrElse
           valLower.StartsWith("thermal_") OrElse valLower.StartsWith("protein_") Then
            Return False
        End If

        ' Palabras reservadas y tokens técnicos del motor Protein
        Dim palabrasReservadas As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "protein", "asset", "material", "materials", "appearance", "appearances",
            "physical", "structural", "thermal", "rendermaterial", "rendermaterialschema",
            "commonschema", "assettableofcontents", "instanceproperties", "streamversiontable",
            "schemadefinition", "autodesk", "adsk", "version", "identity", "category",
            "description", "thumbnail", "definition", "instance", "properties", "property",
            "opaque", "transparent", "transparent_color", "metal", "plastic", "wood", "glass",
            "generic", "generic_diffuse", "schema", "base", "root", "color", "transparency",
            "true", "false", "system", "inventor", "standard", "default", "none", "null"
        }
        If palabrasReservadas.Contains(valLower) Then Return False

        ' Descartar si son solo números, signos o puntuación
        If Regex.IsMatch(val, "^[\d\.\-_/\\:;,#]+$") Then Return False

        ' Debe contener al menos una letra válida
        If Not Regex.IsMatch(val, "[a-zA-ZáéíóúÁÉÍÓÚñÑ]") Then Return False

        Return True
    End Function

    Private Function CompararCodigoReglas(rutaAnt As String, rutaNue As String) As List(Of String)
        Dim diff As New List(Of String)
        Dim lineasAnt = System.IO.File.ReadAllLines(rutaAnt)
        Dim lineasNue = System.IO.File.ReadAllLines(rutaNue)

        Dim subsAnt = ExtraerNombresSubrutinas(lineasAnt)
        Dim subsNue = ExtraerNombresSubrutinas(lineasNue)

        Dim subsAgregadas As New List(Of String)
        For Each s In subsNue
            If Not subsAnt.Contains(s) Then subsAgregadas.Add(s)
        Next

        Dim subsEliminadas As New List(Of String)
        For Each s In subsAnt
            If Not subsNue.Contains(s) Then subsEliminadas.Add(s)
        Next

        diff.Add($"Inspección de Reglas iLogic: {Path.GetFileName(rutaNue)}")
        diff.Add($"• Total líneas: {lineasAnt.Length} (anterior) -> {lineasNue.Length} (nuevo)")

        If subsAgregadas.Count > 0 Then
            diff.Add($"• Subrutinas / Funciones agregadas ({subsAgregadas.Count}):")
            For Each s In subsAgregadas : diff.Add($"    [+] {s}") : Next
        End If

        If subsEliminadas.Count > 0 Then
            diff.Add($"• Subrutinas / Funciones eliminadas ({subsEliminadas.Count}):")
            For Each s In subsEliminadas : diff.Add($"    [-] {s}") : Next
        End If

        If subsAgregadas.Count = 0 AndAlso subsEliminadas.Count = 0 Then
            diff.Add("• Lógica interna de subrutinas existente modificada.")
        End If

        Return diff
    End Function

    Private Function ExtraerNombresSubrutinas(lineas As String()) As HashSet(Of String)
        Dim subs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim reg = New Regex("^\s*(?:Public\s+|Private\s+|Friend\s+)?(Sub|Function)\s+([a-zA-Z0-9_]+)", RegexOptions.IgnoreCase)

        For Each l In lineas
            Dim m = reg.Match(l)
            If m.Success Then
                subs.Add($"{m.Groups(1).Value} {m.Groups(2).Value}")
            End If
        Next

        Return subs
    End Function

    Private Function CompararArchivosXml(rutaAnt As String, rutaNue As String) As List(Of String)
        Dim diff As New List(Of String)
        Dim fiAnt As New FileInfo(rutaAnt)
        Dim fiNue As New FileInfo(rutaNue)

        diff.Add($"Inspección de Configuración/Estilos XML: {Path.GetFileName(rutaNue)}")
        diff.Add($"• Plantilla de estilos actualizada ({FormatBytes(fiAnt.Length)} -> {FormatBytes(fiNue.Length)}).")
        Return diff
    End Function

    Public Sub GuardarRegistroLog(serverRoot As String, localRoot As String, operacion As String, contenidoLog As String)
        Try
            Dim timestamp = DateTime.Now.ToString("yyyy-MM-dd_HHmmss")
            Dim nombreLogIndividual = $"Sync_Log_{timestamp}_{operacion}_{SysEnv.UserName}.txt"

            Dim headerGeneral As New StringBuilder()
            headerGeneral.AppendLine("================================================================================")
            headerGeneral.AppendLine($"RULEMANAGER V2 - REPORTE DE SINCRONIZACIÓN DE BIBLIOTECAS [{operacion}]")
            headerGeneral.AppendLine($"Fecha/Hora: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            headerGeneral.AppendLine($"Usuario: {SysEnv.UserName} ({SysEnv.MachineName})")
            headerGeneral.AppendLine($"Carpeta Local: {localRoot}")
            headerGeneral.AppendLine($"Carpeta Servidor: {serverRoot}")
            headerGeneral.AppendLine("================================================================================")
            headerGeneral.AppendLine()
            headerGeneral.Append(contenidoLog)
            headerGeneral.AppendLine()

            Dim textoFinal = headerGeneral.ToString()

            If Directory.Exists(serverRoot) Then
                Try
                    Dim serverLogsDir = System.IO.Path.Combine(serverRoot, "Sync_Logs")
                    If Not Directory.Exists(serverLogsDir) Then Directory.CreateDirectory(serverLogsDir)

                    Dim rutaLogIndividual = System.IO.Path.Combine(serverLogsDir, nombreLogIndividual)
                    System.IO.File.WriteAllText(rutaLogIndividual, textoFinal, Encoding.UTF8)

                    Dim rutaLogGeneral = System.IO.Path.Combine(serverLogsDir, "CHANGELOG_GENERAL.txt")
                    System.IO.File.AppendAllText(rutaLogGeneral, textoFinal & vbCrLf & vbCrLf, Encoding.UTF8)
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"No se pudo escribir log en servidor: {ex.Message}")
                End Try
            End If

            If Directory.Exists(localRoot) Then
                Try
                    Dim localLogsDir = System.IO.Path.Combine(localRoot, "Sync_Logs")
                    If Not Directory.Exists(localLogsDir) Then Directory.CreateDirectory(localLogsDir)

                    Dim rutaLogLocal = System.IO.Path.Combine(localLogsDir, nombreLogIndividual)
                    System.IO.File.WriteAllText(rutaLogLocal, textoFinal, Encoding.UTF8)
                Catch
                End Try
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error al guardar logs: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Obtiene el propietario del archivo bajo demanda con caché concurrente.
    ''' </summary>
    Public Function ObtenerPropietarioLazy(ruta As String) As String
        If String.IsNullOrWhiteSpace(ruta) OrElse Not File.Exists(ruta) Then Return "-"

        Dim factory As Func(Of String, String) =
            Function(r As String) As String
                Try
                    Dim fi As New FileInfo(r)
                    Dim ac = fi.GetAccessControl()
                    Dim owner = ac.GetOwner(GetType(NTAccount))

                    If owner IsNot Nothing Then
                        Dim nombre = owner.Value

                        If nombre.Contains("\") Then
                            Return nombre.Split("\"c)(1)
                        End If

                        Return nombre
                    End If
                Catch
                End Try

                Return SysEnv.UserName
            End Function

        Return _cachePropietarios.GetOrAdd(ruta, factory)
    End Function

    Private Function GetRelativePath(basePath As String, fullPath As String) As String
        If Not basePath.EndsWith("\") Then basePath &= "\"
        If fullPath.StartsWith(basePath, StringComparison.OrdinalIgnoreCase) Then
            Return fullPath.Substring(basePath.Length)
        End If
        Return Path.GetFileName(fullPath)
    End Function

    Public Shared Function FormatBytes(bytes As Long) As String
        If Math.Abs(bytes) >= 1024 * 1024 Then
            Return $"{bytes / (1024.0 * 1024.0):F2} MB"
        ElseIf Math.Abs(bytes) >= 1024 Then
            Return $"{bytes / 1024.0:F1} KB"
        Else
            Return $"{bytes} B"
        End If
    End Function

End Class
