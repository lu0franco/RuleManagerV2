Imports System.Drawing
Imports System.Windows.Forms

Public Class IconManager

    ' Clase interna que hereda de AxHost para acceder a GetIPictureDispFromPicture
    Private Class PictureDispHost
        Inherits AxHost

        Public Sub New()
            MyBase.New("00000000-0000-0000-0000-000000000000")
        End Sub

        Public Shared Function GetPictureDisp(image As Image) As Object
            Return AxHost.GetIPictureDispFromPicture(image)
        End Function

    End Class

    Public Shared Function ToPictureDisp(ByVal image As Image) As Object

        If image Is Nothing Then Return Nothing

        Return PictureDispHost.GetPictureDisp(image)

    End Function

End Class