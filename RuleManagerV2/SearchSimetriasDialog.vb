Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Imports Color = System.Drawing.Color
Imports Point = System.Drawing.Point

''' <summary>
''' Diálogo interactivo para revisión, edición y confirmación de piezas y ensamblajes simétricos detectados.
''' Permite al usuario verificar el método de detección, auditar el Part Number original sugerido y
''' aplicar los cambios de manera segura y transparente.
''' </summary>
Public Class SearchSimetriasDialog
    Inherits Form

    <DllImport("user32.dll")>
    Private Shared Function EnableWindow(hWnd As IntPtr, bEnable As Boolean) As Boolean
    End Function

    Private ReadOnly _invApp As Inventor.Application
    Private ReadOnly _asmDoc As AssemblyDocument
    Private ReadOnly _engine As SearchSimetriasEngine
    Private ReadOnly _items As List(Of SimetriaItemInfo)

    ' Controles UI
    Private _dgv As DataGridView
    Private _lblResumen As Label
    Private _lblTitulo As Label
    Private _chkSeleccionarTodos As CheckBox
    Private _btnAplicar As Button
    Private _btnCancelar As Button
    Private _btnResaltar As Button
    Private _btnExportar As Button

    Public Sub New(invApp As Inventor.Application,
                   asmDoc As AssemblyDocument,
                   engine As SearchSimetriasEngine,
                   items As List(Of SimetriaItemInfo))

        _invApp = invApp
        _asmDoc = asmDoc
        _engine = engine
        _items = items

        InitializeComponentManual()
        CargarDatos()
    End Sub

    Private Sub InitializeComponentManual()
        Me.Text = "Detector de Simetrías - RuleManager"
        Me.Size = New Size(950, 560)
        Me.MinimumSize = New Size(750, 400)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.BackColor = Color.FromArgb(245, 247, 250)
        Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)

        ' Panel Superior
        Dim pnlHeader As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 70,
            .BackColor = Color.FromArgb(30, 41, 59),
            .Padding = New Padding(16, 12, 16, 12)
        }

        _lblTitulo = New Label With {
            .Text = "Simetrías Detectadas en el Ensamblaje",
            .Font = New Font("Segoe UI", 12.0F, FontStyle.Bold),
            .ForeColor = Color.White,
            .AutoSize = True,
            .Location = New Point(16, 12)
        }

        _lblResumen = New Label With {
            .Text = "Se analizaron las ocurrencias y se encontraron componentes espejo.",
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular),
            .ForeColor = Color.FromArgb(203, 213, 225),
            .AutoSize = True,
            .Location = New Point(17, 38)
        }

        pnlHeader.Controls.Add(_lblTitulo)
        pnlHeader.Controls.Add(_lblResumen)

        ' Panel de Acciones Superiores
        Dim pnlToolbar As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 40,
            .BackColor = Color.FromArgb(241, 245, 249),
            .Padding = New Padding(12, 6, 12, 6)
        }

        _chkSeleccionarTodos = New CheckBox With {
            .Text = "Marcar / Desmarcar todos",
            .Checked = True,
            .AutoSize = True,
            .Location = New Point(16, 10),
            .Cursor = Cursors.Hand
        }
        AddHandler _chkSeleccionarTodos.CheckedChanged, AddressOf ChkSeleccionarTodos_CheckedChanged

        _btnResaltar = New Button With {
            .Text = "Resaltar Ocurrencia en Inventor",
            .Size = New Size(210, 26),
            .Location = New Point(230, 6),
            .BackColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        _btnResaltar.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        AddHandler _btnResaltar.Click, AddressOf BtnResaltar_Click

        _btnExportar = New Button With {
            .Text = "Exportar Lista (.txt)",
            .Size = New Size(140, 26),
            .Location = New Point(450, 6),
            .BackColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        _btnExportar.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225)
        AddHandler _btnExportar.Click, AddressOf BtnExportar_Click

        pnlToolbar.Controls.Add(_chkSeleccionarTodos)
        pnlToolbar.Controls.Add(_btnResaltar)
        pnlToolbar.Controls.Add(_btnExportar)

        ' DataGridView Central
        _dgv = New DataGridView With {
            .Dock = DockStyle.Fill,
            .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.None,
            .RowHeadersVisible = False,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .MultiSelect = False,
            .AutoGenerateColumns = False,
            .GridColor = Color.FromArgb(226, 232, 240)
        }

        _dgv.EnableHeadersVisualStyles = False
        _dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(241, 245, 249)
        _dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(51, 65, 85)
        _dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
        _dgv.ColumnHeadersHeight = 32

        Dim colAplicar As New DataGridViewCheckBoxColumn With {
            .Name = "colAplicar",
            .HeaderText = "Aplicar",
            .Width = 60,
            .DataPropertyName = "AplicarCambio"
        }
        Dim colOcurrencia As New DataGridViewTextBoxColumn With {
            .Name = "colOcurrencia",
            .HeaderText = "Ocurrencia",
            .Width = 170,
            .ReadOnly = True,
            .DataPropertyName = "OccurrenceName"
        }
        Dim colMetodo As New DataGridViewTextBoxColumn With {
            .Name = "colMetodo",
            .HeaderText = "Método Detección",
            .Width = 160,
            .ReadOnly = True,
            .DataPropertyName = "MetodoDeteccionTexto"
        }
        Dim colPNActual As New DataGridViewTextBoxColumn With {
            .Name = "colPNActual",
            .HeaderText = "PN Actual",
            .Width = 130,
            .ReadOnly = True,
            .DataPropertyName = "PartNumberActual"
        }
        Dim colPNNuevo As New DataGridViewTextBoxColumn With {
            .Name = "colPNNuevo",
            .HeaderText = "PN Original Asignar",
            .Width = 150,
            .DataPropertyName = "PartNumberOriginalSugerido"
        }
        Dim colDescNueva As New DataGridViewTextBoxColumn With {
            .Name = "colDescNueva",
            .HeaderText = "Descripción",
            .Width = 110,
            .DataPropertyName = "DescripcionNueva"
        }
        Dim colDetalle As New DataGridViewTextBoxColumn With {
            .Name = "colDetalle",
            .HeaderText = "Detalle Técnico",
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            .ReadOnly = True,
            .DataPropertyName = "DetalleMetodo"
        }

        _dgv.Columns.AddRange(New DataGridViewColumn() {
            colAplicar, colOcurrencia, colMetodo, colPNActual, colPNNuevo, colDescNueva, colDetalle
        })

        ' Panel Inferior de Botones
        Dim pnlFooter As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 54,
            .BackColor = Color.FromArgb(241, 245, 249),
            .Padding = New Padding(16, 10, 16, 10)
        }

        _btnAplicar = New Button With {
            .Text = "Aplicar Cambios",
            .DialogResult = DialogResult.OK,
            .Size = New Size(140, 34),
            .BackColor = Color.FromArgb(16, 185, 129),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand,
            .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
        }
        _btnAplicar.FlatAppearance.BorderSize = 0
        _btnAplicar.Location = New Point(pnlFooter.Width - 300, 10)

        _btnCancelar = New Button With {
            .Text = "Cancelar",
            .DialogResult = DialogResult.Cancel,
            .Size = New Size(110, 34),
            .BackColor = Color.FromArgb(226, 232, 240),
            .ForeColor = Color.FromArgb(71, 85, 105),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular),
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand,
            .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
        }
        _btnCancelar.FlatAppearance.BorderSize = 0
        _btnCancelar.Location = New Point(pnlFooter.Width - 130, 10)

        pnlFooter.Controls.Add(_btnAplicar)
        pnlFooter.Controls.Add(_btnCancelar)

        Me.Controls.Add(_dgv)
        Me.Controls.Add(pnlToolbar)
        Me.Controls.Add(pnlHeader)
        Me.Controls.Add(pnlFooter)

        Me.AcceptButton = _btnAplicar
        Me.CancelButton = _btnCancelar
    End Sub

    Private Sub CargarDatos()
        _dgv.DataSource = Nothing
        _dgv.DataSource = _items

        _lblResumen.Text = $"Se detectaron {_items.Count} piezas/ensamblajes como simetrías en '{_asmDoc.DisplayName}'."
    End Sub

    Private Sub ChkSeleccionarTodos_CheckedChanged(sender As Object, e As EventArgs)
        Dim estado As Boolean = _chkSeleccionarTodos.Checked
        For Each item As SimetriaItemInfo In _items
            item.AplicarCambio = estado
        Next
        _dgv.Refresh()
    End Sub

    Private Sub BtnResaltar_Click(sender As Object, e As EventArgs)
        Try
            If _dgv.CurrentRow Is Nothing Then Return
            Dim item As SimetriaItemInfo = DirectCast(_dgv.CurrentRow.DataBoundItem, SimetriaItemInfo)
            If item IsNot Nothing AndAlso item.Occurrence IsNot Nothing Then
                _asmDoc.SelectSet.Clear()
                _asmDoc.SelectSet.Select(item.Occurrence)
            End If
        Catch ex As Exception
            MessageBox.Show("No se pudo seleccionar la ocurrencia: " & ex.Message, "Resaltar", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub BtnExportar_Click(sender As Object, e As EventArgs)
        Try
            Dim sfd As New SaveFileDialog With {
                .Filter = "Archivo de texto (*.txt)|*.txt|Archivo CSV (*.csv)|*.csv",
                .FileName = "Reporte_Simetrias_" & DateTime.Now.ToString("yyyyMMdd_HHmmss")
            }

            If sfd.ShowDialog() = DialogResult.OK Then
                Using sw As New StreamWriter(sfd.FileName)
                    sw.WriteLine("OCURRENCIA;METODO_DETECCION;PART_NUMBER_ACTUAL;PART_NUMBER_ORIGINAL;DESCRIPCION;DETALLE")
                    For Each item As SimetriaItemInfo In _items
                        sw.WriteLine($"{item.OccurrenceName};{item.MetodoDeteccionTexto};{item.PartNumberActual};{item.PartNumberOriginalSugerido};{item.DescripcionNueva};{item.DetalleMetodo}")
                    Next
                End Using
                MessageBox.Show("Reporte exportado exitosamente en:" & vbCrLf & sfd.FileName,
                                "Exportar", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Error al exportar reporte: " & ex.Message, "Error Exportar", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Garantiza que la ventana de Inventor no quede bloqueada cuando se muestra el formulario.
    ''' </summary>
    Public Sub HabilitarVentanaInventor()
        Try
            Dim hwndInv As IntPtr = New IntPtr(_invApp.MainFrameHWND)
            If hwndInv <> IntPtr.Zero Then
                EnableWindow(hwndInv, True)
            End If
        Catch
        End Try
    End Sub

End Class
