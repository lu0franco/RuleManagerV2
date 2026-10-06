Imports System.Windows.Forms
Imports IO = System.IO

Public Class FolderPickerDialog

    Private _selectedPath As String = ""
    Private _treeView As TreeView
    Private _rootPath As String
    Private _initialPath As String

    Public ReadOnly Property SelectedPath As String
        Get
            Return _selectedPath
        End Get
    End Property

    Public Sub New(rootPath As String, Optional initialPath As String = "")
        _rootPath = rootPath
        _initialPath = initialPath
    End Sub

    Public Function ShowDialog(owner As IWin32Window) As DialogResult
        Dim form As New Form()
        form.Text = "Seleccionar carpeta destino"
        form.Size = New Drawing.Size(450, 550)
        form.StartPosition = FormStartPosition.CenterParent

        _treeView = New TreeView()
        _treeView.Dock = DockStyle.Fill
        _treeView.Font = New Drawing.Font("Segoe UI", 9.5F)
        _treeView.FullRowSelect = True
        _treeView.HideSelection = False
        _treeView.ShowLines = True
        _treeView.ShowPlusMinus = True

        BuildTree()

        Dim panelBottom As New Panel()
        panelBottom.Dock = DockStyle.Bottom
        panelBottom.Height = 50
        panelBottom.Padding = New Padding(10)

        Dim btnOk As New Button()
        btnOk.Text = "Aceptar"
        btnOk.DialogResult = DialogResult.OK
        btnOk.Size = New Drawing.Size(80, 28)
        btnOk.Location = New Drawing.Point(240, 10)

        Dim btnCancel As New Button()
        btnCancel.Text = "Cancelar"
        btnCancel.DialogResult = DialogResult.Cancel
        btnCancel.Size = New Drawing.Size(80, 28)
        btnCancel.Location = New Drawing.Point(330, 10)

        panelBottom.Controls.Add(btnOk)
        panelBottom.Controls.Add(btnCancel)

        form.Controls.Add(_treeView)
        form.Controls.Add(panelBottom)

        AddHandler _treeView.AfterSelect, Sub(s, e)
                                              _selectedPath = CStr(e.Node.Tag)
                                          End Sub

        AddHandler btnOk.Click, Sub(s, e)
                                    If String.IsNullOrEmpty(_selectedPath) Then
                                        MessageBox.Show("Seleccioná una carpeta.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                        Dim f = CType(s, Button).FindForm()
                                        f.DialogResult = DialogResult.None
                                    End If
                                End Sub

        Return form.ShowDialog(owner)
    End Function

    Private Sub BuildTree()
        Dim rootNode As New TreeNode(IO.Path.GetFileName(_rootPath))
        rootNode.Tag = _rootPath

        Dim carpetasBase As New List(Of String) From {
            "01 - DIBUJOS", "02 - PLANOS", "03 - CORTES",
            "04 - COMERCIALES", "05 - PIEZAS COMUNES",
            "06 - MULTIMEDIA", "07 - AUTOCAD - SAT",
            "08 - MAQUINAS DE REFERENCIAS"
        }

        For Each carpeta In carpetasBase
            Dim carpetaPath = IO.Path.Combine(_rootPath, carpeta)
            Dim node As New TreeNode(carpeta)
            node.Tag = carpetaPath

            If IO.Directory.Exists(carpetaPath) Then
                AddSubDirectories(node, carpetaPath, 3)
            End If

            rootNode.Nodes.Add(node)
        Next

        _treeView.Nodes.Add(rootNode)
        rootNode.Expand()

        ' Seleccionar nodo inicial si existe
        If Not String.IsNullOrEmpty(_initialPath) Then
            SelectNodeByPath(rootNode, _initialPath)
        End If
    End Sub

    Private Sub AddSubDirectories(parentNode As TreeNode, path As String, maxDepth As Integer)
        If maxDepth <= 0 Then Return
        Try
            ' FIX BC30068: renombrar Dir -> subDir para evitar conflicto con la función incorporada Dir()
            For Each subDir In IO.Directory.GetDirectories(path)
                Dim dirName = IO.Path.GetFileName(subDir)
                If dirName.Equals("OldVersions", StringComparison.OrdinalIgnoreCase) Then Continue For

                Dim node As New TreeNode(dirName)
                node.Tag = subDir
                AddSubDirectories(node, subDir, maxDepth - 1)
                parentNode.Nodes.Add(node)
            Next
        Catch
        End Try
    End Sub

    Private Sub SelectNodeByPath(parentNode As TreeNode, targetPath As String)
        For Each node As TreeNode In parentNode.Nodes
            If String.Equals(CStr(node.Tag), targetPath, StringComparison.OrdinalIgnoreCase) Then
                _treeView.SelectedNode = node
                node.EnsureVisible()
                Return
            End If
            SelectNodeByPath(node, targetPath)
        Next
    End Sub

End Class