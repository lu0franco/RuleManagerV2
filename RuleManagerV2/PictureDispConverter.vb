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

End Class