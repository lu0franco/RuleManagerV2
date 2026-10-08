Imports System.Windows.Forms
Imports System.Drawing

Public Class PictureDispConverter
    Inherits AxHost

    Private Sub New()
        MyBase.New("")
    End Sub

    Public Shared Function ToIPictureDisp(
        image As Image) As Object

        Return GetIPictureDispFromPicture(image)

    End Function

    Public Shared Function ToImage(
        pictureDisp As Object) As Image

        Try
            If pictureDisp IsNot Nothing Then
                Return GetPictureFromIPicture(pictureDisp)
            End If
        Catch
        End Try
        Return Nothing

    End Function

End Class