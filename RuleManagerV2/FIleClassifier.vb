Imports Inventor
Imports IO = System.IO

Public Class FileClassifier

    ' =========================================================
    ' CARPETAS A IGNORAR (EN CUALQUIER NIVEL DE LA RUTA)
    ' =========================================================
    Private Shared ReadOnly _carpetasExcluidas As New List(Of String) From {
        "OldVersions"
    }

    ' =========================================================
    ' CARPETAS ESPECIALES QUE VAN A 01 - DIBUJOS CON SUBCARPETAS
    ' =========================================================
    Private Shared ReadOnly _carpetasEspeciales As New List(Of String) From {
        "AIP",
        "Frame"
    }

    ' =========================================================
    ' REGLAS DE USUARIO (extensión -> carpeta destino)
    ' =========================================================
    Private Shared _userRules As New Dictionary(Of String, String)

    ' =========================================================
    ' RUTA DEL LOG
    ' =========================================================
    Private Shared ReadOnly _logPath As String = "C:\Temp\Reorder.log"

    ' =========================================================
    ' LOG
    ' =========================================================
    Public Shared Sub WriteLog(message As String)
        Try
            Dim dir As String = IO.Path.GetDirectoryName(_logPath)
            If Not IO.Directory.Exists(dir) Then
                IO.Directory.CreateDirectory(dir)
            End If

            If IO.File.Exists(_logPath) Then
                Dim fi As New IO.FileInfo(_logPath)
                If fi.IsReadOnly Then
                    fi.IsReadOnly = False
                End If
            End If

            Using sw As New IO.StreamWriter(_logPath, True)
                sw.WriteLine($"{Now:yyyy-MM-dd HH:mm:ss} | {message}")
            End Using
        Catch
        End Try
    End Sub

    ' =========================================================
    ' GUARDAR/CARGAR REGLAS DE USUARIO
    ' =========================================================
    Public Shared Sub SaveUserRule(ext As String, folderName As String)
        _userRules(ext.ToLower()) = folderName
        WriteLog($"REGLA USUARIO: {ext} -> {folderName}")
    End Sub

    Public Shared Function GetUserRule(ext As String) As String
        Dim key As String = ext.ToLower()
        If _userRules.ContainsKey(key) Then
            Return _userRules(key)
        End If
        Return ""
    End Function

    ' =========================================================
    ' VERIFICAR EXCLUSIONES EN TODA LA RUTA RELATIVA
    ' =========================================================
    Private Shared Function IsInExcludedPath(filePath As String, sourceRoot As String) As Boolean
        If String.IsNullOrEmpty(sourceRoot) Then
            Dim parentFolder As String = IO.Path.GetFileName(IO.Path.GetDirectoryName(filePath))
            For Each carpetaExcluida As String In _carpetasExcluidas
                If parentFolder.Equals(carpetaExcluida, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next
            Return False
        End If

        Dim root As String = sourceRoot
        If Not root.EndsWith("\") Then root &= "\"

        If Not filePath.StartsWith(root, StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If

        If root.Length > filePath.Length Then
            Return False
        End If

        Dim relative As String = filePath.Substring(root.Length)
        If String.IsNullOrEmpty(relative) Then
            Return False
        End If

        Dim parts As String() = relative.Split("\"c)

        For Each part In parts
            For Each carpetaExcluida As String In _carpetasExcluidas
                If part.Equals(carpetaExcluida, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next
        Next

        Return False
    End Function

    ' =========================================================
    ' DETECTAR SI ESTÁ DENTRO DE CARPETA ESPECIAL (AIP/Frame)
    ' =========================================================
    Private Shared Function IsInSpecialFolder(filePath As String, sourceRoot As String) As Boolean
        If String.IsNullOrEmpty(sourceRoot) Then
            Dim parentFolder As String = IO.Path.GetFileName(IO.Path.GetDirectoryName(filePath))
            For Each carpetaEspecial As String In _carpetasEspeciales
                If parentFolder.Equals(carpetaEspecial, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next
            Return False
        End If

        Dim root As String = sourceRoot
        If Not root.EndsWith("\") Then root &= "\"

        If Not filePath.StartsWith(root, StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If

        If root.Length > filePath.Length Then
            Return False
        End If

        Dim relative As String = filePath.Substring(root.Length)
        If String.IsNullOrEmpty(relative) Then
            Return False
        End If

        Dim parts As String() = relative.Split("\"c)

        For Each part In parts
            For Each carpetaEspecial As String In _carpetasEspeciales
                If part.Equals(carpetaEspecial, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next
        Next

        Return False
    End Function

    ' =========================================================
    ' LIMPIAR "Nueva carpeta" DE LA RUTA RELATIVA
    ' =========================================================
    Private Shared Function CleanRelativePath(relativeFolderPath As String) As String
        If String.IsNullOrEmpty(relativeFolderPath) Then Return ""

        Dim parts As List(Of String) = relativeFolderPath.Split("\"c).ToList()
        Dim cleaned As New List(Of String)

        For Each part In parts
            If Not part.ToLower().Contains("nueva carpeta") Then
                cleaned.Add(part)
            End If
        Next

        Return String.Join("\", cleaned)
    End Function

    ' =========================================================
    ' OBTENER SOLO LA CARPETA BASE DE DESTINO
    ' =========================================================
    Public Shared Function GetTargetFolder(
        invApp As Inventor.Application,
        filePath As String,
        Optional sourceRoot As String = "") As String

        If IsInExcludedPath(filePath, sourceRoot) Then
            WriteLog($"EXCLUIDO (carpeta prohibida): {filePath}")
            Return ""
        End If

        Dim ext As String = IO.Path.GetExtension(filePath).ToLower()
        Dim fileName As String = IO.Path.GetFileNameWithoutExtension(filePath).ToUpper()
        Dim estaEnCarpetaEspecial As Boolean = IsInSpecialFolder(filePath, sourceRoot)

        ' Verificar reglas de usuario primero
        Dim userRule As String = GetUserRule(ext)
        If Not String.IsNullOrEmpty(userRule) Then
            Return userRule
        End If

        Select Case ext

            Case ".ipt", ".iam"
                If estaEnCarpetaEspecial Then
                    Return "01 - DIBUJOS"
                End If
                If HasComCheck(invApp, filePath) Then
                    Return "04 - COMERCIALES"
                End If
                Return "01 - DIBUJOS"

            Case ".idw", ".ipn"
                Return "02 - PLANOS"

            Case ".dwg", ".dxf"
                If fileName.Contains("MULTICORTE") Then
                    Return "03 - CORTES"
                End If
                Return "02 - PLANOS"

            ' --- CAD NEUTRO ---
            Case ".step", ".stp", ".sat", ".x_b", ".x_t"
                Return "07 - AUTOCAD - SAT"

            ' --- DOCUMENTACION ---
            Case ".doc", ".docx", ".xlsx", ".pdf"
                Return "02 - PLANOS"

            ' --- ACCESOS DIRECTOS (maquinas de referencia) ---
            Case ".lnk"
                Return "08 - MAQUINAS DE REFERENCIAS"

            ' --- IPJ (archivo de proyecto) ---
            Case ".ipj"
                Return ""

        End Select

        ' No clasificado → pedir al usuario
        Return "__UNCLASSIFIED__"
    End Function

    ' =========================================================
    ' OBTENER RUTA COMPLETA DE DESTINO (CON SUBCARPETAS)
    ' =========================================================
    Public Shared Function GetDestinationPath(
        invApp As Inventor.Application,
        filePath As String,
        sourceRoot As String,
        targetRoot As String) As String

        Dim baseFolder As String = GetTargetFolder(invApp, filePath, sourceRoot)

        ' Si está excluido o es IPJ, devolver vacío
        If String.IsNullOrEmpty(baseFolder) Then
            Return ""
        End If

        ' Si no está clasificado, devolver marcador especial
        If baseFolder = "__UNCLASSIFIED__" Then
            Return "__UNCLASSIFIED__"
        End If

        Dim relativeSubfolder As String = ""
        If Not String.IsNullOrEmpty(sourceRoot) Then
            Dim root As String = sourceRoot
            If Not root.EndsWith("\") Then root &= "\"

            If filePath.StartsWith(root, StringComparison.OrdinalIgnoreCase) Then
                Dim relativeFull As String = filePath.Substring(root.Length)
                relativeSubfolder = IO.Path.GetDirectoryName(relativeFull)
                If relativeSubfolder Is Nothing Then relativeSubfolder = ""
            End If
        End If

        If Not String.IsNullOrEmpty(relativeSubfolder) Then
            If relativeSubfolder.StartsWith(baseFolder & "\", StringComparison.OrdinalIgnoreCase) Then
                relativeSubfolder = relativeSubfolder.Substring(baseFolder.Length + 1)
            ElseIf String.Equals(relativeSubfolder, baseFolder, StringComparison.OrdinalIgnoreCase) Then
                relativeSubfolder = ""
            End If
        End If

        relativeSubfolder = CleanRelativePath(relativeSubfolder)

        Dim finalDir As String = IO.Path.Combine(targetRoot, baseFolder, relativeSubfolder)
        Dim finalPath As String = IO.Path.Combine(finalDir, IO.Path.GetFileName(filePath))

        WriteLog($"REUBICAR: {filePath} -> {finalPath}")

        Return finalPath
    End Function

    ' =========================================================
    ' HAS COM CHECK
    ' =========================================================
    Private Shared Function HasComCheck(
        invApp As Inventor.Application,
        filePath As String) As Boolean

        Dim doc As Document = Nothing
        Dim yaEstabaAbierto As Boolean = False

        Try
            If Not IO.File.Exists(filePath) Then
                Return False
            End If

            For Each d As Document In invApp.Documents
                If d.FullFileName.Equals(filePath, StringComparison.OrdinalIgnoreCase) Then
                    doc = d
                    yaEstabaAbierto = True
                    Exit For
                End If
            Next

            If doc Is Nothing Then
                doc = invApp.Documents.Open(filePath, False)
            End If

            If doc Is Nothing Then
                Return False
            End If

            Dim props As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")

            For Each p As Inventor.Property In props
                If p.Name = "COM" AndAlso p.Value.ToString().Trim() = "✓" Then
                    If Not yaEstabaAbierto Then doc.Close(True)
                    Return True
                End If
            Next

            For Each p As Inventor.Property In props
                If p.Name = "BULONERIA" AndAlso p.Value.ToString().Trim() = "✓" Then
                    If Not yaEstabaAbierto Then doc.Close(True)
                    Return True
                End If
            Next

            If Not yaEstabaAbierto Then doc.Close(True)
            Return False

        Catch ex As Exception
            WriteLog($"ERROR HasComCheck [{filePath}]: {ex.Message}")
            If doc IsNot Nothing AndAlso Not yaEstabaAbierto Then
                Try
                    doc.Close(True)
                Catch
                End Try
            End If
            Return False
        End Try
    End Function

End Class