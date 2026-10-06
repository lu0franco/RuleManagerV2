Imports System.Drawing
Imports System.Windows.Forms


Public Class OptionsDialog
    Inherits Form

    Private _chkCrearCarpetas As CheckBox
    Private _btnOk As Button
    Private _btnCancel As Button


    ' =========================================================
    ' CONSTRUCTOR - UI construida por codigo (sin designer)
    ' =========================================================
    Public Sub New()

        ' --- Ventana ---
        Text = "RuleManager - Opciones"
        StartPosition = FormStartPosition.CenterScreen
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False
        MinimizeBox = False
        ShowInTaskbar = False
        ClientSize = New Size(420, 170)
        Font = New Font("Segoe UI", 9.0F)

        ' --- Titulo / descripcion ---
        Dim lblTitulo As New Label With {
            .Text = "Comportamiento del addin al cambiar de proyecto:",
            .AutoSize = True,
            .Location = New Point(15, 15)
        }

        ' --- CheckBox: creacion de carpetas ---
        _chkCrearCarpetas = New CheckBox With {
            .Text = "Crear estructura de carpetas del proyecto" & vbCrLf &
                    "(01 - DIBUJOS, 02 - PLANOS, 03 - CORTES, etc.)",
            .Location = New Point(30, 45),
            .Size = New Size(370, 45),
            .Checked = AddinSettings.Current.CrearCarpetasProyecto
        }

        ' --- Boton Aceptar ---
        _btnOk = New Button With {
            .Text = "Aceptar",
            .DialogResult = DialogResult.None,
            .Location = New Point(240, 115),
            .Size = New Size(85, 28)
        }
        AddHandler _btnOk.Click, AddressOf OnOk

        ' --- Boton Cancelar ---
        _btnCancel = New Button With {
            .Text = "Cancelar",
            .DialogResult = DialogResult.Cancel,
            .Location = New Point(330, 115),
            .Size = New Size(75, 28)
        }

        ' --- Agregar controles ---
        Controls.Add(lblTitulo)
        Controls.Add(_chkCrearCarpetas)
        Controls.Add(_btnOk)
        Controls.Add(_btnCancel)

        AcceptButton = _btnOk
        CancelButton = _btnCancel

    End Sub


    ' =========================================================
    ' ACEPTAR: guardar configuracion
    ' =========================================================
    Private Sub OnOk(sender As Object, e As EventArgs)

        AddinSettings.Current.CrearCarpetasProyecto =
            _chkCrearCarpetas.Checked

        AddinSettings.Save()

        DialogResult = DialogResult.OK
        Close()

    End Sub

End Class
