Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports Inventor

' Alias explícitos para evitar colisiones entre System.Windows.Forms / System.Drawing e Inventor
Imports Color = System.Drawing.Color
Imports Point = System.Drawing.Point
Imports Font = System.Drawing.Font
Imports Size = System.Drawing.Size
Imports FontStyle = System.Drawing.FontStyle
Imports TextBox = System.Windows.Forms.TextBox
Imports Button = System.Windows.Forms.Button
Imports Label = System.Windows.Forms.Label
Imports Panel = System.Windows.Forms.Panel
Imports Form = System.Windows.Forms.Form
Imports GroupBox = System.Windows.Forms.GroupBox
Imports InvProperty = Inventor.Property
Imports InvPropertySet = Inventor.PropertySet

''' <summary>
''' Diálogo interactivo para visualizar y editar rápidamente las iProperties del documento activo en Autodesk Inventor.
''' Enfocado exclusivamente en los 5 campos clave de Proyecto:
''' - Nº de pieza (Part Number)
''' - Nº de almacenamiento (Stock Number)
''' - Proyecto (Project)
''' - Diseñador (Designer)
''' - Proveedor / Cliente (Vendor)
''' </summary>
Public Class CargarDatosDialog
    Inherits Form

    Private ReadOnly _invApp As Inventor.Application
    Private ReadOnly _doc As Inventor.Document

    ' Controles - Información de Cabecera
    Private _lblArchivoNombre As Label
    Private _lblArchivoRuta As Label

    ' Controles - Los 5 Campos Principales de Proyecto
    Private _txtPartNumber As TextBox
    Private _txtStockNumber As TextBox
    Private _txtProject As TextBox
    Private _txtDesigner As TextBox
    Private _txtVendor As TextBox

    ' Checkbox
    Private _chkActualizarVistas As System.Windows.Forms.CheckBox

    ' Botones
    Private _btnGuardarYAceptar As Button
    Private _btnAplicar As Button
    Private _btnRecargar As Button
    Private _btnCancelar As Button

    Public Sub New(invApp As Inventor.Application)
        _invApp = invApp
        If _invApp IsNot Nothing Then
            Try
                _doc = _invApp.ActiveDocument
            Catch
                _doc = Nothing
            End Try
        End If

        InitializeComponent()
        CargarDatosDesdeDocumento()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "RuleManager V2 - Cargar Datos (iProperties de Proyecto)"
        Me.Size = New Size(710, 420)
        Me.MinimumSize = New Size(690, 400)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Font = New Font("Segoe UI", 9.0F)
        Me.BackColor = Color.FromArgb(248, 250, 252)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False

        ' Panel de Encabezado Superior
        Dim pnlHeader As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 78,
            .BackColor = Color.FromArgb(15, 23, 42),
            .Padding = New Padding(18, 10, 18, 10)
        }

        ' Sin emojis unicode que puedan causar caracteres no reconocidos '[]' en WinForms
        Dim lblTitulo As New Label With {
            .Text = "CARGAR DATOS DE PROYECTO (iProperties)",
            .Font = New Font("Segoe UI", 11.0F, FontStyle.Bold),
            .ForeColor = Color.White,
            .Location = New Point(18, 8),
            .AutoSize = True
        }

        _lblArchivoNombre = New Label With {
            .Text = "Documento: " & If(_doc IsNot Nothing, _doc.DisplayName, "Sin documento activo"),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(56, 189, 248),
            .Location = New Point(18, 32),
            .AutoSize = True
        }

        _lblArchivoRuta = New Label With {
            .Text = If(_doc IsNot Nothing AndAlso Not String.IsNullOrEmpty(_doc.FullFileName), _doc.FullFileName, "Archivo nuevo no guardado en disco"),
            .Font = New Font("Segoe UI", 8.0F, FontStyle.Regular),
            .ForeColor = Color.FromArgb(148, 163, 184),
            .Location = New Point(18, 52),
            .AutoSize = True,
            .MaximumSize = New Size(650, 18)
        }

        pnlHeader.Controls.AddRange({lblTitulo, _lblArchivoNombre, _lblArchivoRuta})

        ' Panel de Botones Inferior (Footer)
        Dim pnlFooter As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 56,
            .BackColor = Color.FromArgb(241, 245, 249),
            .Padding = New Padding(16, 10, 16, 10)
        }

        _chkActualizarVistas = New System.Windows.Forms.CheckBox With {
            .Text = "Actualizar vistas y LDM tras guardar",
            .Checked = True,
            .Location = New Point(18, 18),
            .AutoSize = True,
            .ForeColor = Color.FromArgb(51, 65, 85)
        }

        _btnGuardarYAceptar = New Button With {
            .Text = "Guardar y Aceptar",
            .Size = New Size(140, 32),
            .Location = New Point(270, 12),
            .BackColor = Color.FromArgb(16, 185, 129),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .Cursor = Cursors.Hand
        }
        _btnGuardarYAceptar.FlatAppearance.BorderSize = 0
        AddHandler _btnGuardarYAceptar.Click, AddressOf OnGuardarYAceptarClicked

        _btnAplicar = New Button With {
            .Text = "Aplicar",
            .Size = New Size(80, 32),
            .Location = New Point(418, 12),
            .BackColor = Color.FromArgb(226, 232, 240),
            .ForeColor = Color.FromArgb(30, 41, 59),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9.0F),
            .Cursor = Cursors.Hand
        }
        _btnAplicar.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        AddHandler _btnAplicar.Click, AddressOf OnAplicarClicked

        _btnRecargar = New Button With {
            .Text = "Recargar",
            .Size = New Size(85, 32),
            .Location = New Point(506, 12),
            .BackColor = Color.FromArgb(226, 232, 240),
            .ForeColor = Color.FromArgb(30, 41, 59),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9.0F),
            .Cursor = Cursors.Hand
        }
        _btnRecargar.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        AddHandler _btnRecargar.Click, Sub() CargarDatosDesdeDocumento()

        _btnCancelar = New Button With {
            .Text = "Cancelar",
            .Size = New Size(80, 32),
            .Location = New Point(599, 12),
            .BackColor = Color.FromArgb(241, 245, 249),
            .ForeColor = Color.FromArgb(100, 116, 139),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9.0F),
            .Cursor = Cursors.Hand
        }
        _btnCancelar.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        AddHandler _btnCancelar.Click, Sub() Me.Close()

        pnlFooter.Controls.AddRange({_chkActualizarVistas, _btnGuardarYAceptar, _btnAplicar, _btnRecargar, _btnCancelar})

        ' Panel de Contenido Central
        Dim pnlContent As New Panel With {
            .Dock = DockStyle.Fill,
            .AutoScroll = False,
            .Padding = New Padding(16, 12, 16, 12)
        }

        ' =========================================================================
        ' GRUPO ÚNICO: DATOS PRINCIPALES DE PROYECTO
        ' (Se retiró por completo la sección de Información Complementaria a solicitud del usuario)
        ' =========================================================================
        Dim grpPrincipales As New GroupBox With {
            .Text = " DATOS DE PROYECTO ",
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(3, 105, 161),
            .Location = New Point(16, 10),
            .Size = New Size(662, 225),
            .BackColor = Color.White
        }

        ' Ancho ampliado a 160 px para evitar que 'Nº de almacenamiento:' se corte
        Dim lblPartNumber As New Label With {
            .Text = "Nº de pieza:",
            .Location = New Point(20, 30),
            .Size = New Size(155, 20),
            .ForeColor = Color.FromArgb(30, 41, 59),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
        }
        _txtPartNumber = New TextBox With {
            .Location = New Point(180, 27),
            .Size = New Size(455, 24),
            .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        }

        Dim lblStockNumber As New Label With {
            .Text = "Nº de almacenamiento:",
            .Location = New Point(20, 68),
            .Size = New Size(155, 20),
            .ForeColor = Color.FromArgb(30, 41, 59),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
        }
        _txtStockNumber = New TextBox With {
            .Location = New Point(180, 65),
            .Size = New Size(455, 24),
            .Font = New Font("Segoe UI", 9.5F, FontStyle.Regular)
        }

        Dim lblProject As New Label With {
            .Text = "Proyecto (OT):",
            .Location = New Point(20, 106),
            .Size = New Size(155, 20),
            .ForeColor = Color.FromArgb(30, 41, 59),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
        }
        _txtProject = New TextBox With {
            .Location = New Point(180, 103),
            .Size = New Size(215, 24),
            .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        }

        Dim lblDesigner As New Label With {
            .Text = "Diseñador:",
            .Location = New Point(415, 106),
            .Size = New Size(75, 20),
            .ForeColor = Color.FromArgb(30, 41, 59),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
        }
        _txtDesigner = New TextBox With {
            .Location = New Point(495, 103),
            .Size = New Size(140, 24),
            .Font = New Font("Segoe UI", 9.5F, FontStyle.Regular)
        }

        Dim lblVendor As New Label With {
            .Text = "Proveedor / Cliente:",
            .Location = New Point(20, 144),
            .Size = New Size(155, 20),
            .ForeColor = Color.FromArgb(30, 41, 59),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
        }
        _txtVendor = New TextBox With {
            .Location = New Point(180, 141),
            .Size = New Size(455, 24),
            .Font = New Font("Segoe UI", 9.5F, FontStyle.Regular)
        }

        Dim lblTip As New Label With {
            .Text = "Nota: Estos campos alimentan las Listas de Materiales, Carátulas y Planos en la carpeta del proyecto.",
            .Location = New Point(20, 185),
            .Size = New Size(615, 25),
            .Font = New Font("Segoe UI", 8.0F, FontStyle.Italic),
            .ForeColor = Color.FromArgb(100, 116, 139)
        }

        grpPrincipales.Controls.AddRange({
            lblPartNumber, _txtPartNumber,
            lblStockNumber, _txtStockNumber,
            lblProject, _txtProject,
            lblDesigner, _txtDesigner,
            lblVendor, _txtVendor,
            lblTip
        })

        pnlContent.Controls.Add(grpPrincipales)

        Me.Controls.AddRange({pnlContent, pnlFooter, pnlHeader})

        ' Acceso rápido con teclado
        Me.AcceptButton = _btnGuardarYAceptar
        Me.CancelButton = _btnCancelar
    End Sub

    ''' <summary>
    ''' Carga los valores actuales de las iProperties desde el documento activo de Autodesk Inventor.
    ''' </summary>
    Private Sub CargarDatosDesdeDocumento()
        If _doc Is Nothing Then
            MessageBox.Show("No se encontró ningún documento activo en Autodesk Inventor.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Try
            Dim dtProps As InvPropertySet = Nothing
            Try
                dtProps = _doc.PropertySets.Item("Design Tracking Properties")
            Catch
            End Try

            If dtProps IsNot Nothing Then
                _txtPartNumber.Text = ObtenerValorPropiedad(dtProps, "Part Number")
                _txtStockNumber.Text = ObtenerValorPropiedad(dtProps, "Stock Number")
                _txtProject.Text = ObtenerValorPropiedad(dtProps, "Project")
                _txtDesigner.Text = ObtenerValorPropiedad(dtProps, "Designer")
                _txtVendor.Text = ObtenerValorPropiedad(dtProps, "Vendor")
            End If

            ' Si Part Number estuviera vacío, pre-cargar nombre de archivo sin extensión como sugerencia
            If String.IsNullOrWhiteSpace(_txtPartNumber.Text) Then
                _txtPartNumber.Text = System.IO.Path.GetFileNameWithoutExtension(_doc.FullFileName)
            End If

        Catch ex As Exception
            MessageBox.Show("Error al leer iProperties del documento: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Guarda los cambios en las iProperties dentro de una transacción reversible de Autodesk Inventor.
    ''' </summary>
    Private Function GuardarDatosEnDocumento() As Boolean
        If _doc Is Nothing Then
            MessageBox.Show("No hay documento activo para guardar.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        Dim trans As Transaction = Nothing
        Try
            If _invApp IsNot Nothing Then
                trans = _invApp.TransactionManager.StartTransaction(_doc, "RuleManager - Cargar Datos iProperties")
            End If

            Dim dtProps As InvPropertySet = Nothing
            Try
                dtProps = _doc.PropertySets.Item("Design Tracking Properties")
            Catch
                dtProps = _doc.PropertySets.Add("Design Tracking Properties")
            End Try

            AsignarValorPropiedad(dtProps, "Part Number", _txtPartNumber.Text.Trim())
            AsignarValorPropiedad(dtProps, "Stock Number", _txtStockNumber.Text.Trim())
            AsignarValorPropiedad(dtProps, "Project", _txtProject.Text.Trim())
            AsignarValorPropiedad(dtProps, "Designer", _txtDesigner.Text.Trim())
            AsignarValorPropiedad(dtProps, "Vendor", _txtVendor.Text.Trim())

            If trans IsNot Nothing Then
                trans.End()
                trans = Nothing
            End If

            If _chkActualizarVistas.Checked Then
                Try
                    _doc.Update()
                Catch
                End Try
            End If

            Return True
        Catch ex As Exception
            If trans IsNot Nothing Then
                Try
                    trans.Abort()
                Catch
                End Try
            End If
            MessageBox.Show("Error al guardar iProperties: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Sub OnGuardarYAceptarClicked(sender As Object, e As EventArgs)
        If GuardarDatosEnDocumento() Then
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End If
    End Sub

    Private Sub OnAplicarClicked(sender As Object, e As EventArgs)
        If GuardarDatosEnDocumento() Then
            _btnAplicar.Text = "Aplicado"
            Dim timer As New System.Windows.Forms.Timer With {.Interval = 1500}
            AddHandler timer.Tick, Sub()
                                       _btnAplicar.Text = "Aplicar"
                                       timer.Stop()
                                       timer.Dispose()
                                   End Sub
            timer.Start()
        End If
    End Sub

    Private Function ObtenerValorPropiedad(pSet As InvPropertySet, nombreProp As String) As String
        If pSet Is Nothing Then Return ""
        Try
            For Each p As InvProperty In pSet
                If String.Equals(p.Name, nombreProp, StringComparison.OrdinalIgnoreCase) Then
                    Return If(p.Value IsNot Nothing, p.Value.ToString(), "")
                End If
            Next
        Catch
        End Try
        Return ""
    End Function

    Private Sub AsignarValorPropiedad(pSet As InvPropertySet, nombreProp As String, nuevoValor As String)
        If pSet Is Nothing Then Exit Sub
        Try
            Dim propEncontrada As InvProperty = Nothing
            For Each p As InvProperty In pSet
                If String.Equals(p.Name, nombreProp, StringComparison.OrdinalIgnoreCase) Then
                    propEncontrada = p
                    Exit For
                End If
            Next

            If propEncontrada IsNot Nothing Then
                propEncontrada.Value = nuevoValor
            Else
                pSet.Add(nuevoValor, nombreProp)
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error al asignar {nombreProp}: {ex.Message}")
        End Try
    End Sub

End Class
