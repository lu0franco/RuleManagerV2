Imports System.IO
Imports System.Linq

Imports IOPath = System.IO.Path
Imports IOFile = System.IO.File
Imports IODirectory = System.IO.Directory

''' <summary>
''' Renombra físicamente los archivos .idw en una carpeta de proyecto.
''' Se ejecuta DESPUÉS de que ProjectRenamer cerró todos los documentos,
''' para evitar locks de Inventor sobre los archivos.
''' 
''' Formato de renombrado:
''' - IDW más pesado: {NombreProyecto} - 1.idw
''' - Segundo más pesado: {NombreProyecto} - 2.idw
''' - Tercero más pesado: {NombreProyecto} - 3.idw
''' - etc.
''' </summary>
Public Class IdwPhysicalRenamer

    Private ReadOnly _logPath As String

    Public Sub New()
        _logPath = "C:\Temp\ProjectRenamer.log"
    End Sub

    ''' <summary>
    ''' Renombra los IDWs de un proyecto según el nombre del proyecto y su peso.
    ''' </summary>
    ''' <param name="targetFolder">Carpeta raíz del proyecto nuevo</param>
    ''' <param name="newProjectName">Nombre del proyecto nuevo</param>
    ''' <param name="oldAsmName">Nombre del IAM viejo que se reemplazó (ej: "Ensamblaje sin fin elevador.iam")</param>
    Public Sub Execute(targetFolder As String, newProjectName As String, oldAsmName As String)

        Log("===== IDW PHYSICAL RENAME START =====")

        If Not IODirectory.Exists(targetFolder) Then
            Throw New DirectoryNotFoundException("La carpeta destino no existe: " & targetFolder)
        End If

        ' Buscar todos los IDWs (excluyendo OldVersions)
        Dim allIdws = IODirectory.GetFiles(targetFolder, "*.idw", SearchOption.AllDirectories) _
            .Where(Function(f) Not f.ToLower().Contains("\oldversions")) _
            .Select(Function(f) New FileInfo(f)) _
            .OrderByDescending(Function(fi) fi.Length) _
            .ToList()

        Log("IDWs encontrados para renombrar: " & allIdws.Count)

        ' =========================================================
        ' RENOMBRAR CON SUFIJO NUMÉRICO SEGÚN PESO
        ' =========================================================
        ' #1 = más pesado, #2 = segundo más pesado, etc.
        For i As Integer = 0 To allIdws.Count - 1

            Dim idwInfo As FileInfo = allIdws(i)
            Dim oldIdwPath As String = idwInfo.FullName
            Dim oldBase As String = IOPath.GetFileNameWithoutExtension(oldIdwPath)

            ' Construir nuevo nombre: {Proyecto} - {número}.idw
            Dim suffix As Integer = i + 1
            Dim newIdwName As String = newProjectName & " - " & suffix.ToString() & ".idw"

            Dim targetIdwPath As String = IOPath.Combine(IOPath.GetDirectoryName(oldIdwPath), newIdwName)

            ' Si ya tiene el nombre correcto, saltear
            If String.Equals(oldIdwPath, targetIdwPath, StringComparison.OrdinalIgnoreCase) Then
                Log("-> [SKIP] Ya tiene el nombre correcto: " & oldBase)
                Continue For
            End If

            Log("-> Renombrando: " & IOPath.GetFileName(oldIdwPath) & " -> " & newIdwName)

            ' ✅ INTENTAR CON MÁS REINTENTOS Y MAYOR DELAY
            Dim success As Boolean = False
            For retry As Integer = 0 To 30
                Try
                    If IOFile.Exists(targetIdwPath) Then
                        IOFile.Delete(targetIdwPath)
                    End If

                    IOFile.Move(oldIdwPath, targetIdwPath)
                    Log("-> OK renombrado (intento " & (retry + 1) & ")")
                    success = True
                    Exit For

                Catch ex As IOException
                    If retry = 30 Then
                        Log("-> ERROR: No se pudo renombrar después de 31 intentos: " & ex.Message)
                        Throw
                    End If
                    ' ✅ DELAY AUMENTADO entre reintentos
                    System.Threading.Thread.Sleep(1000)
                End Try
            Next

            If Not success Then
                Log("-> [FALLO] No se pudo renombrar: " & oldBase)
            End If

        Next

        Log("===== IDW PHYSICAL RENAME END =====")

    End Sub

    Private Sub Log(msg As String)
        Try
            IOFile.AppendAllText(_logPath, DateTime.Now.ToString("HH:mm:ss") & " | [PHYSICAL] " & msg & vbCrLf)
        Catch
            ' Si no puede escribir el log, no fallar todo
        End Try
    End Sub

End Class