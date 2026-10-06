Imports System.Windows.Forms

Public Class WindowWrapper
    Implements IWin32Window

    Private _hwnd As IntPtr

    Public Sub New(hwnd As IntPtr)
        _hwnd = hwnd
    End Sub

    Public ReadOnly Property Handle As IntPtr Implements IWin32Window.Handle
        Get
            Return _hwnd
        End Get
    End Property
End Class