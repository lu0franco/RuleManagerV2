Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Windows.Forms
Imports Inventor
Imports IO = System.IO

''' <summary>
''' Gestor de sincronización y ajuste de cantidades para Listas de Piezas (PartsList) en dibujos IDW / DWG.
''' 
''' Problema que resuelve:
''' Cuando una pieza (ej. "PLATINA DE ANCLAJE" o "PLATINA PATA") se utiliza en varios subensamblajes
''' (ej. 4 veces en CHASIS INFERIOR [8 por fabricarse 2] y 4 veces en CHASIS SUPERIOR [8 por fabricarse 2]),
''' la iProperty 3D 'CANTIDAD_USADA' vale 16 (correcto para el plano de fabricación individual de esa pieza).
''' Sin embargo, en el plano del subensamblaje 'CHASIS INFERIOR', la columna CANT de la lista de piezas
''' mostraría 16 en lugar de las 8 requeridas para ese subconjunto.
''' 
''' Solución implementada:
''' 1. Detecta automáticamente la columna de cantidad (título 'CANT', 'CANT.', 'CANTIDAD', 'CANTIDAD_USADA', etc.).
''' 2. Identifica el subensamblaje 3D referenciado por la vista o lista de piezas.
''' 3. Cuenta las ocurrencias reales de cada pieza dentro de dicho subensamblaje.
''' 4. Aplica una sobreescritura estática exclusiva a nivel de celda en la tabla (cell.Static = True, cell.Value = nuevaCantidad),
'''    garantizando que el plano muestre exactamente la cantidad necesaria para ese subconjunto
'''    SIN alterar la iProperty 3D 'CANTIDAD_USADA' del modelo para que su plano de fabricación siga marcando el total general.
''' </summary>
Public Class PartsListQuantityManager

    Private ReadOnly _invApp As Inventor.Application
    Private ReadOnly _log As New List(Of String)

    Private Const LOG_PATH As String = "C:\Temp\PartsListQuantity.log"

    Public Sub New(invApp As Inventor.Application)
        _invApp = invApp
    End Sub

    ''' <summary>
    ''' Detecta la columna que representa la cantidad en la lista de piezas.
    ''' Compatible con estilos estándar de la empresa donde la columna se titula 'CANT' y mapea a CANTIDAD_USADA.
    ''' </summary>
    Public Shared Function DetectarColumnaCantidad(pl As PartsList) As Integer
        If pl Is Nothing OrElse pl.PartsListColumns.Count = 0 Then Return -1

        ' 1. Prioridad: Buscar columna titulada explícitamente "CANT" o "CANT."
        For cIdx As Integer = 1 To pl.PartsListColumns.Count
            Dim col As PartsListColumn = pl.PartsListColumns.Item(cIdx)
            Dim titleUpper As String = col.Title.Trim().ToUpper()
            If titleUpper = "CANT" OrElse titleUpper = "CANT." Then
                Return cIdx
            End If
        Next

        ' 2. Segunda prioridad: Buscar "CANTIDAD", "CANTIDAD_USADA", "CANT. USADA"
        For cIdx As Integer = 1 To pl.PartsListColumns.Count
            Dim col As PartsListColumn = pl.PartsListColumns.Item(cIdx)
            Dim titleUpper As String = col.Title.Trim().ToUpper()
            If titleUpper = "CANTIDAD" OrElse
               titleUpper = "CANTIDAD_USADA" OrElse
               titleUpper.Contains("CANTIDAD USADA") OrElse
               titleUpper.Contains("CANT. USADA") OrElse
               titleUpper = "CANTIDAD DE ELEMENTOS" Then
                Return cIdx
            End If
        Next

        ' 3. Tercera prioridad: Buscar "QTY" o "ITEM QTY"
        For cIdx As Integer = 1 To pl.PartsListColumns.Count
            Dim col As PartsListColumn = pl.PartsListColumns.Item(cIdx)
            Dim titleUpper As String = col.Title.Trim().ToUpper()
            If titleUpper = "QTY" OrElse titleUpper = "ITEM QTY" OrElse titleUpper.Contains("QTY") Then
                Return cIdx
            End If
        Next

        ' 4. Cuarta prioridad: Si la segunda columna suele ser la de cantidad (columna 2 en estilo estándar)
        If pl.PartsListColumns.Count >= 2 Then
            Return 2
        End If

        Return 1
    End Function

    ''' <summary>
    ''' Obtiene el documento de ensamblaje (.iam) referenciado por la lista de piezas o sus vistas asociadas.
    ''' </summary>
    Public Shared Function ObtenerEnsamblajeReferenciado(pl As PartsList, invApp As Inventor.Application) As AssemblyDocument
        If pl Is Nothing Then Return Nothing

        ' 1. A través de la DrawingView asociada a la PartsList
        Try
            If pl.DrawingView IsNot Nothing AndAlso pl.DrawingView.ReferencedDocumentDescriptor IsNot Nothing Then
                Dim refDoc As Document = pl.DrawingView.ReferencedDocumentDescriptor.ReferencedDocument
                If refDoc IsNot Nothing AndAlso refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Return DirectCast(refDoc, AssemblyDocument)
                End If
            End If
        Catch
        End Try

        ' 2. A través de ReferencedDocumentDescriptor de la PartsList
        Try
            If pl.ReferencedDocumentDescriptor IsNot Nothing Then
                Dim refDoc As Document = pl.ReferencedDocumentDescriptor.ReferencedDocument
                If refDoc IsNot Nothing AndAlso refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Return DirectCast(refDoc, AssemblyDocument)
                End If
            End If
        Catch
        End Try

        ' 3. A través de ReferencedFile de la PartsList
        Try
            If pl.ReferencedFile IsNot Nothing Then
                Dim filePath As String = pl.ReferencedFile.FullFileName
                If filePath.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then
                    For Each d As Document In invApp.Documents
                        If String.Equals(d.FullFileName, filePath, StringComparison.OrdinalIgnoreCase) AndAlso
                           d.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                            Return DirectCast(d, AssemblyDocument)
                        End If
                    Next
                    Dim openedDoc As Document = invApp.Documents.Open(filePath, False)
                    If openedDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Return DirectCast(openedDoc, AssemblyDocument)
                    End If
                End If
            End If
        Catch
        End Try

        ' 4. Buscar en las vistas de la hoja que contiene la PartsList
        Try
            Dim sheet As Sheet = DirectCast(pl.Parent, Sheet)
            For Each v As DrawingView In sheet.DrawingViews
                If v.ReferencedDocumentDescriptor IsNot Nothing Then
                    Dim refDoc As Document = v.ReferencedDocumentDescriptor.ReferencedDocument
                    If refDoc IsNot Nothing AndAlso refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Return DirectCast(refDoc, AssemblyDocument)
                    End If
                End If
            Next
        Catch
        End Try

        ' 5. Buscar en las demás hojas del dibujo
        Try
            Dim drawDoc As DrawingDocument = DirectCast(pl.Parent.Parent, DrawingDocument)
            For Each s As Sheet In drawDoc.Sheets
                For Each v As DrawingView In s.DrawingViews
                    If v.ReferencedDocumentDescriptor IsNot Nothing Then
                        Dim refDoc As Document = v.ReferencedDocumentDescriptor.ReferencedDocument
                        If refDoc IsNot Nothing AndAlso refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                            Return DirectCast(refDoc, AssemblyDocument)
                        End If
                    End If
                Next
            Next
        Catch
        End Try

        Return Nothing
    End Function

    ''' <summary>
    ''' Ejecuta el flujo principal abriendo el diálogo interactivo para revisión y edición manual o automática.
    ''' </summary>
    Public Function Ejecutar() As Boolean
        Try
            _log.Clear()
            Log("=== INICIO AJUSTE DE CANTIDADES EN LISTAS DE PIEZAS (IDW) ===")

            If _invApp.ActiveDocument Is Nothing OrElse
               _invApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                MessageBox.Show("Esta herramienta solo funciona en documentos de plano o dibujo (.idw / .dwg).",
                                "Ajustar Cantidades LDM", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False
            End If

            Dim drawDoc As DrawingDocument = DirectCast(_invApp.ActiveDocument, DrawingDocument)
            Dim sheet As Sheet = drawDoc.ActiveSheet

            Dim targetPartsList As PartsList = Nothing
            If sheet.PartsLists.Count > 0 Then
                targetPartsList = sheet.PartsLists.Item(1)
            Else
                For Each s As Sheet In drawDoc.Sheets
                    If s.PartsLists.Count > 0 Then
                        targetPartsList = s.PartsLists.Item(1)
                        sheet = s
                        Exit For
                    End If
                Next
            End If

            If targetPartsList Is Nothing Then
                MessageBox.Show("No se encontraron Listas de Piezas (PartsList) en el plano actual.",
                                "Ajustar Cantidades LDM", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return False
            End If

            Dim colCantIdx As Integer = DetectarColumnaCantidad(targetPartsList)
            If colCantIdx = -1 Then
                MessageBox.Show("No se encontró la columna de cantidad (CANT / CANTIDAD) en la lista de piezas.",
                                "Ajustar Cantidades LDM", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False
            End If

            Dim asmDoc As AssemblyDocument = ObtenerEnsamblajeReferenciado(targetPartsList, _invApp)

            Using dlg As New AjustarCantidadesDialog(_invApp, targetPartsList, sheet, asmDoc, colCantIdx)
                Dim hwnd As IntPtr = IntPtr.Zero
                Try
                    hwnd = New IntPtr(_invApp.MainFrameHWND)
                Catch
                End Try

                Dim res As DialogResult
                If hwnd <> IntPtr.Zero Then
                    res = dlg.ShowDialog(New WindowWrapper(hwnd))
                Else
                    res = dlg.ShowDialog()
                End If

                Return (res = DialogResult.OK)
            End Using

        Catch ex As Exception
            Dim err As String = "Error al ajustar cantidades en lista de piezas: " & ex.Message
            Log(err)
            MessageBox.Show(err, "Error - Ajustar Cantidades LDM", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Permite restaurar las cantidades dinámicas originales vinculadas al modelo (quita el override estático).
    ''' </summary>
    Public Function RestaurarValoresOriginales() As Boolean
        Try
            Dim drawDoc As DrawingDocument = DirectCast(_invApp.ActiveDocument, DrawingDocument)
            Dim totalRestauradas As Integer = 0

            For sheetIndex As Integer = 1 To drawDoc.Sheets.Count
                Dim sheet As Sheet = drawDoc.Sheets.Item(sheetIndex)
                For plIndex As Integer = 1 To sheet.PartsLists.Count
                    Dim pl As PartsList = sheet.PartsLists.Item(plIndex)
                    Dim colCantIndex As Integer = DetectarColumnaCantidad(pl)

                    If colCantIndex > 0 Then
                        For rIdx As Integer = 1 To pl.PartsListRows.Count
                            Dim cell As PartsListCell = pl.PartsListRows.Item(rIdx).Item(colCantIndex)
                            If cell.Static Then
                                cell.Static = False
                                totalRestauradas += 1
                            End If
                        Next
                    End If
                Next
            Next

            drawDoc.Update2(True)
            _invApp.ActiveView.Update()

            MessageBox.Show($"Se restauraron {totalRestauradas} celdas al valor dinámico original del modelo 3D.",
                            "Restaurar Cantidades Originales", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return True

        Catch ex As Exception
            MessageBox.Show("Error al restaurar valores: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Sub Log(mensaje As String)
        _log.Add(mensaje)
        Try
            Dim dir As String = IO.Path.GetDirectoryName(LOG_PATH)
            If Not IO.Directory.Exists(dir) Then IO.Directory.CreateDirectory(dir)
            Using sw As New IO.StreamWriter(LOG_PATH, True)
                sw.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {mensaje}")
            End Using
        Catch
        End Try
    End Sub

End Class
