Imports System.Windows.Forms
Imports System.Drawing

Public Class ProjectRenameDialog
    Inherits Form

    ' =========================================================
    ' CONTROLES
    ' =========================================================
    Private txtModelo As TextBox
    Private txtMaquina As TextBox
    Private txtOT As TextBox
    Private txtCliente As TextBox

    Private lblSep1 As Label
    Private lblSep2 As Label
    Private lblSep3 As Label

    Private btnAceptar As Button
    Private btnCancelar As Button

    Private lblTitulo As Label
    Private lblPreview As Label
    Private lblPreviewTitle As Label

    ' =========================================================
    ' DATOS INICIALES (guardados antes de crear controles)
    ' =========================================================
    Private _dataInicial As ProjectData

    ' =========================================================
    ' PROPIEDADES PÚBLICAS
    ' =========================================================
    Public ReadOnly Property ResultData As ProjectData
        Get
            ' ✅ Protección: si los controles no existen, devolver datos iniciales
            If txtModelo Is Nothing Then
                Return _dataInicial
            End If

            Return New ProjectData() With {
                .Modelo = txtModelo.Text.Trim(),
                .Maquina = txtMaquina.Text.Trim(),
                .OT = txtOT.Text.Trim(),
                .Cliente = txtCliente.Text.Trim()
            }
        End Get
    End Property

    ' =========================================================
    ' CONSTANTES DE DISEÑO
    ' =========================================================
    Private Const PADDING_PIXELS As Integer = 8
    Private Const MIN_WIDTH As Integer = 55
    Private Const MAX_WIDTH As Integer = 260
    Private Const GAP As Integer = 8
    Private Const SEPARATOR_WIDTH As Integer = 14
    Private Const LABEL_HEIGHT As Integer = 16
    Private Const TEXTBOX_HEIGHT As Integer = 26
    Private Const ROW_GAP As Integer = 4

    ' =========================================================
    ' CONSTRUCTOR
    ' =========================================================
    Public Sub New(data As ProjectData)

        ' ✅ Guardar datos primero (antes de crear controles)
        _dataInicial = New ProjectData() With {
            .Modelo = data.Modelo,
            .Maquina = data.Maquina,
            .OT = data.OT,
            .Cliente = data.Cliente
        }

        ' ✅ Crear controles primero
        InitializeComponent()

        ' ✅ Asignar valores a los controles YA CREADOS
        txtModelo.Text = _dataInicial.Modelo
        txtMaquina.Text = _dataInicial.Maquina
        txtOT.Text = _dataInicial.OT
        txtCliente.Text = _dataInicial.Cliente

        ' ✅ Actualizar preview con los valores asignados
        UpdatePreview()

    End Sub

    ' =========================================================
    ' INICIALIZACIÓN DE COMPONENTES
    ' =========================================================
    Private Sub InitializeComponent()

        Me.Text = "Confirmar Nombre del Proyecto"
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = System.Drawing.Color.White

        ' =========================================
        ' TITULO
        ' =========================================
        lblTitulo = New Label()
        lblTitulo.Text = "Verifique los datos del proyecto. Modifique si es necesario:"
        lblTitulo.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        lblTitulo.Location = New System.Drawing.Point(20, 15)
        lblTitulo.AutoSize = True
        Me.Controls.Add(lblTitulo)

        ' =========================================
        ' ANCHOS DINAMICOS (usar _dataInicial, NO ResultData)
        ' =========================================
        Dim wModelo As Integer = CalculateTextBoxWidth(_dataInicial.Modelo)
        Dim wMaquina As Integer = CalculateTextBoxWidth(_dataInicial.Maquina)
        Dim wOT As Integer = CalculateTextBoxWidth(_dataInicial.OT)
        Dim wCliente As Integer = CalculateTextBoxWidth(_dataInicial.Cliente)

        ' =========================================
        ' POSICIONES
        ' =========================================
        Dim xModelo As Integer = 20
        Dim xSep1 As Integer = xModelo + wModelo + GAP
        Dim xMaquina As Integer = xSep1 + SEPARATOR_WIDTH + GAP
        Dim xSep2 As Integer = xMaquina + wMaquina + GAP
        Dim xOT As Integer = xSep2 + SEPARATOR_WIDTH + GAP
        Dim xSep3 As Integer = xOT + wOT + GAP
        Dim xCliente As Integer = xSep3 + SEPARATOR_WIDTH + GAP

        Dim totalWidth As Integer = xCliente + wCliente + 40

        Me.Size = New Size(
            Math.Max(totalWidth, 500),
            260
        )

        ' =========================================
        ' LABELS DE CAMPOS
        ' =========================================
        Dim labelY As Integer = 55

        Me.Controls.Add(CreateLabel("Modelo", xModelo, labelY, wModelo))
        Me.Controls.Add(CreateLabel("Máquina", xMaquina, labelY, wMaquina))
        Me.Controls.Add(CreateLabel("OT", xOT, labelY, wOT))
        Me.Controls.Add(CreateLabel("Cliente", xCliente, labelY, wCliente))

        ' =========================================
        ' TEXTBOXES
        ' =========================================
        Dim textY As Integer = labelY + LABEL_HEIGHT + ROW_GAP

        txtModelo = CreateTextBox(xModelo, textY, wModelo)
        Me.Controls.Add(txtModelo)

        lblSep1 = CreateSeparator(xSep1, textY)
        Me.Controls.Add(lblSep1)

        txtMaquina = CreateTextBox(xMaquina, textY, wMaquina)
        Me.Controls.Add(txtMaquina)

        lblSep2 = CreateSeparator(xSep2, textY)
        Me.Controls.Add(lblSep2)

        txtOT = CreateTextBox(xOT, textY, wOT)
        Me.Controls.Add(txtOT)

        lblSep3 = CreateSeparator(xSep3, textY)
        Me.Controls.Add(lblSep3)

        txtCliente = CreateTextBox(xCliente, textY, wCliente)
        Me.Controls.Add(txtCliente)

        ' =========================================
        ' PREVIEW
        ' =========================================
        Dim previewY As Integer = textY + TEXTBOX_HEIGHT + 18

        lblPreviewTitle = New Label()
        lblPreviewTitle.Text = "Nombre resultante:"
        lblPreviewTitle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        lblPreviewTitle.Location = New System.Drawing.Point(20, previewY)
        lblPreviewTitle.AutoSize = True
        Me.Controls.Add(lblPreviewTitle)

        lblPreview = New Label()
        lblPreview.Text = ""
        lblPreview.Font = New Font("Segoe UI", 9, FontStyle.Italic)
        lblPreview.ForeColor = System.Drawing.Color.DarkBlue
        lblPreview.Location = New System.Drawing.Point(20, previewY + 22)
        lblPreview.Size = New Size(totalWidth - 40, 25)
        Me.Controls.Add(lblPreview)

        ' Eventos para actualizar preview en tiempo real
        AddHandler txtModelo.TextChanged, AddressOf OnPreviewTextChanged
        AddHandler txtMaquina.TextChanged, AddressOf OnPreviewTextChanged
        AddHandler txtOT.TextChanged, AddressOf OnPreviewTextChanged
        AddHandler txtCliente.TextChanged, AddressOf OnPreviewTextChanged

        ' =========================================
        ' BOTONES
        ' =========================================
        Dim buttonY As Integer = previewY + 55

        btnAceptar = New Button()
        btnAceptar.Text = "Aceptar y Renombrar"
        btnAceptar.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        btnAceptar.Size = New Size(170, 32)
        btnAceptar.Location = New System.Drawing.Point(20, buttonY)
        btnAceptar.BackColor =System.Drawing.Color.FromArgb(0, 120, 212)
        btnAceptar.ForeColor = System.Drawing.Color.White
        btnAceptar.FlatStyle = FlatStyle.Flat
        btnAceptar.FlatAppearance.BorderSize = 0
        btnAceptar.DialogResult = DialogResult.OK
        Me.Controls.Add(btnAceptar)

        btnCancelar = New Button()
        btnCancelar.Text = "Cancelar"
        btnCancelar.Size = New Size(120, 32)
        btnCancelar.Location = New System.Drawing.Point(200, buttonY)
        btnCancelar.DialogResult = DialogResult.Cancel
        Me.Controls.Add(btnCancelar)

        Me.AcceptButton = btnAceptar
        Me.CancelButton = btnCancelar

    End Sub

    ' =========================================================
    ' EVENTOS
    ' =========================================================
    Private Sub OnPreviewTextChanged(sender As Object, e As EventArgs)
        UpdatePreview()
    End Sub

    ' =========================================================
    ' HELPERS DE CREACIÓN DE CONTROLES
    ' =========================================================
    Private Function CreateLabel(
        text As String,
        x As Integer,
        y As Integer,
        width As Integer) As Label

        Dim lbl As New Label()
        lbl.Text = text
        lbl.Font = New Font("Segoe UI", 8)
        lbl.Location = New System.Drawing.Point(x, y)
        lbl.Size = New Size(width, LABEL_HEIGHT)

        Return lbl

    End Function

    ' ✅ Crear TextBox vacío (el texto se asigna después)
    Private Function CreateTextBox(
        x As Integer,
        y As Integer,
        width As Integer) As TextBox

        Dim txt As New TextBox()
        txt.Text = ""
        txt.Font = New Font("Segoe UI", 10)
        txt.Location = New System.Drawing.Point(x, y)
        txt.Size = New Size(width, TEXTBOX_HEIGHT)
        txt.BorderStyle = BorderStyle.FixedSingle
        txt.TextAlign = HorizontalAlignment.Center

        Return txt

    End Function

    Private Function CreateSeparator(
        x As Integer,
        y As Integer) As Label

        Dim lbl As New Label()
        lbl.Text = "-"
        lbl.Font = New Font("Segoe UI", 11, FontStyle.Bold)
        lbl.Location = New System.Drawing.Point(x, y + 3)
        lbl.Size = New Size(SEPARATOR_WIDTH, TEXTBOX_HEIGHT)
        lbl.TextAlign = ContentAlignment.MiddleCenter

        Return lbl

    End Function

    ' =========================================================
    ' CÁLCULO DE ANCHO DINÁMICO
    ' =========================================================
    Private Function CalculateTextBoxWidth(
        text As String) As Integer

        If String.IsNullOrWhiteSpace(text) Then
            Return MIN_WIDTH
        End If

        Using g As Graphics = Me.CreateGraphics()

            Dim font As New Font("Segoe UI", 10)

            Dim textSize As SizeF =
                g.MeasureString(text, font)

            Dim width As Integer =
                CInt(Math.Ceiling(textSize.Width)) +
                PADDING_PIXELS

            Return Math.Max(
                MIN_WIDTH,
                Math.Min(MAX_WIDTH, width)
            )

        End Using

    End Function

    ' =========================================================
    ' ACTUALIZACIÓN DE PREVIEW
    ' =========================================================
    Private Sub UpdatePreview()
        ' ✅ Protección: si lblPreview no existe aún, salir
        If lblPreview Is Nothing Then
            Exit Sub
        End If

        lblPreview.Text = BuildPreviewName()
    End Sub

    Private Function BuildPreviewName() As String

        ' ✅ Usar los textos de los controles directamente (ya existen)
        Return CleanName(txtModelo.Text) & " - " &
               CleanName(txtMaquina.Text) & " - " &
               CleanName(txtOT.Text) & " - " &
               CleanName(txtCliente.Text)

    End Function

    ' =========================================================
    ' LIMPIEZA DE NOMBRE (consistente con ProjectRenamer)
    ' =========================================================
    Private Function CleanName(
        value As String) As String

        If String.IsNullOrEmpty(value) Then
            Return ""
        End If

        Dim cleanValue As String = value

        For Each c As Char In IO.Path.GetInvalidFileNameChars()
            ' Preservar símbolo de grado °
            If c = ChrW(176) Then Continue For
            cleanValue = cleanValue.Replace(c, "-"c)
        Next

        Return cleanValue.Trim()

    End Function

End Class