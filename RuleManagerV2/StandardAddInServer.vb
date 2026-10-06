Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

<GuidAttribute("1a7a3453-e509-4689-aa1b-e7a2415dcf60"),
 ComVisible(True)>
Public Class StandardAddInServer

    Implements Inventor.ApplicationAddInServer

    Private _buttons As Buttons
    Private _invApp As Inventor.Application

    Public Sub Activate(
        AddInSiteObject As ApplicationAddInSite,
        FirstTime As Boolean) _
        Implements ApplicationAddInServer.Activate

        Try

            _invApp = AddInSiteObject.Application

            _buttons = New Buttons(_invApp)

            StartupTasks.Initialize(_invApp)

        Catch ex As Exception

            MessageBox.Show(ex.ToString())

        End Try

    End Sub


    Public Sub Deactivate() _
        Implements ApplicationAddInServer.Deactivate

        Try

            StartupTasks.Shutdown()

        Catch

        End Try

    End Sub


    Public Sub ExecuteCommand(
        CommandID As Integer) _
        Implements ApplicationAddInServer.ExecuteCommand

    End Sub


    Public ReadOnly Property Automation _
        As Object _
        Implements ApplicationAddInServer.Automation

        Get

            Return Nothing

        End Get

    End Property

End Class