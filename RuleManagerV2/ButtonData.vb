Imports Inventor

Public Class ButtonData

    ' =========================================================
    ' PROPIEDADES
    ' =========================================================

    Public Property DisplayName As String

    Public Property InternalName As String

    Public Property Description As String

    Public Property Environment As String

    Public Property LargeIcon As String

    Public Property ExecuteAction As Action(Of Inventor.Application)


    ' =========================================================
    ' CONSTRUCTOR
    ' =========================================================

    Public Sub New(
        displayName As String,
        internalName As String,
        description As String,
        environment As String,
        executeAction As Action(Of Inventor.Application),
largeIcon As String)

        Me.DisplayName = displayName

        Me.InternalName = internalName

        Me.Description = description

        Me.Environment = environment

        Me.ExecuteAction = executeAction

        Me.LargeIcon = largeIcon

    End Sub

End Class