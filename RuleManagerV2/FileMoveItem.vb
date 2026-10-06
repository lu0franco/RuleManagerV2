Imports IO = System.IO

Public Class FileMoveItem

    Public Property SourcePath As String
    Public Property ProposedDestination As String
    Public Property FinalDestination As String
    Public Property FileName As String
    Public Property Extension As String
    Public Property RelativeSource As String
    Public Property TargetFolder As String
    Public Property IsClassified As Boolean
    Public Property IsExcluded As Boolean
    Public Property Status As String ' "Pendiente", "Movido", "Ignorado", "Error"

    Public Sub New(sourcePath As String, rootPath As String)
        Me.SourcePath = sourcePath
        Me.FileName = IO.Path.GetFileName(sourcePath)
        Me.Extension = IO.Path.GetExtension(sourcePath).ToLower()
        Me.Status = "Pendiente"

        ' Calcular ruta relativa desde root
        If sourcePath.StartsWith(rootPath, StringComparison.OrdinalIgnoreCase) Then
            Me.RelativeSource = sourcePath.Substring(rootPath.Length).TrimStart("\"c)
        Else
            Me.RelativeSource = Me.FileName
        End If
    End Sub

    Public ReadOnly Property IsAlreadyInPlace As Boolean
        Get
            Return String.Equals(SourcePath, FinalDestination, StringComparison.OrdinalIgnoreCase)
        End Get
    End Property

    Public ReadOnly Property DisplayDestination As String
        Get
            If IsExcluded Then Return "(excluido)"
            If String.IsNullOrEmpty(FinalDestination) Then Return "???"
            If IsAlreadyInPlace Then Return "(ya está aquí)"
            Return FinalDestination
        End Get
    End Property

End Class