Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Windows.Forms
Imports Inventor

Imports IOPath = System.IO.Path
Imports IOFile = System.IO.File

''' <summary>
''' Gestor automático para la detección del tipo de componente (Pieza .ipt vs Ensamblaje .iam)
''' e inserción automática de la Lista de Piezas (PartsList) correspondiente según las opciones del Add-In.
''' 
''' Regla de activación:
''' Se dispara cuando se inserta una vista o dos vistas simultáneamente en una hoja de dibujo
''' y la hoja no cuenta aún con una lista de piezas (sheet.PartsLists.Count = 0).
''' Toma la primera vista insertada por índice (DrawingViews.Item(1)).
''' </summary>
Public Class AutoPartsListManager

    Private Shared _invApp As Inventor.Application
    Private Shared _isProcessing As Boolean = False
    Private Shared _debounceTimer As System.Windows.Forms.Timer = Nothing
    Private Shared ReadOnly _timerLock As New Object()
    Private Const LOG_PATH As String = "C:\Temp\RuleManager.log"

    Public Sub New(invApp As Inventor.Application)
        _invApp = invApp
    End Sub

    Public Shared Sub Initialize(invApp As Inventor.Application)
        _invApp = invApp
    End Sub

    ''' <summary>
    ''' Ejecuta la verificación e inserción automática de la lista de piezas en la hoja activa o indicada.
    ''' </summary>
    ''' <param name="sheet">Hoja objetivo (si es Nothing, toma la hoja activa del documento de dibujo).</param>
    ''' <param name="forzar">Si es True, omite las comprobaciones de primera vista y configuración para inserción manual.</param>
    Public Function Ejecutar(Optional sheet As Sheet = Nothing, Optional forzar As Boolean = False) As Boolean
        If _isProcessing Then Return False
        _isProcessing = True

        Try
            If _invApp Is Nothing Then Return False

            Dim drawDoc As DrawingDocument = Nothing
            If _invApp.ActiveDocument IsNot Nothing AndAlso
               _invApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                drawDoc = DirectCast(_invApp.ActiveDocument, DrawingDocument)
            End If

            If drawDoc Is Nothing Then
                Log("Ejecutar: No hay un documento de dibujo activo (.idw / .dwg).")
                Return False
            End If

            Dim targetSheet As Sheet = sheet
            If targetSheet Is Nothing Then
                targetSheet = drawDoc.ActiveSheet
            End If

            If targetSheet Is Nothing Then
                Log("Ejecutar: No se pudo determinar la hoja activa.")
                Return False
            End If

            ' 1. Validar activación en Opciones
            Dim cfg = AddinSettings.Current
            If Not forzar AndAlso Not cfg.AutoInsertarListaPiezas Then
                Log("Inserción automática de lista de piezas deshabilitada en Opciones.")
                Return False
            End If

            ' 2. Regla clave del usuario:
            ' "al insertar una vista en una hoja que NO tenga una vista previa (es decir que la que metemos es la primera)"
            If Not forzar Then
                If targetSheet.DrawingViews.Count <> 1 Then
                    Log($"Hoja '{targetSheet.Name}' tiene {targetSheet.DrawingViews.Count} vistas. No es la primera vista. Se omite.")
                    Return False
                End If

                If targetSheet.PartsLists.Count > 0 Then
                    Log($"Hoja '{targetSheet.Name}' ya posee una Lista de Piezas ({targetSheet.PartsLists.Count}). Se omite.")
                    Return False
                End If
            Else
                If targetSheet.DrawingViews.Count = 0 Then
                    MessageBox.Show("La hoja no contiene ninguna vista para vincular a la lista de piezas.",
                                    "Insertar Lista de Piezas", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Return False
                End If
            End If

            ' 3. Obtener la primera vista de la hoja
            Dim oView As DrawingView = targetSheet.DrawingViews.Item(1)
            If oView Is Nothing Then
                Log("No se pudo obtener la vista de la hoja.")
                Return False
            End If

            ' 4. Determinar si el componente referenciado es una Pieza (.ipt) o un Ensamblaje (.iam)
            Dim tipoDoc As DocumentTypeEnum = DeterminarTipoComponente(oView)
            Dim esPieza As Boolean = (tipoDoc = DocumentTypeEnum.kPartDocumentObject)

            Log($"Vista '{oView.Name}' detectada como: {If(esPieza, "PIEZA (.ipt)", "ENSAMBLAJE (.iam)")}")

            ' 5. Calcular punto de colocación según configuración
            Dim tg As TransientGeometry = _invApp.TransientGeometry
            Dim oPoint As Point2d = CalcularPuntoColocacion(targetSheet, cfg.UbicacionListaPiezas, tg)

            ' 6. Determinar nivel de lista de materiales (BOM Level)
            Dim bomLevel As PartsListLevelEnum = PartsListLevelEnum.kStructured
            If Not esPieza Then
                If cfg.NivelBomEnsamblaje?.Equals("SoloPiezas", StringComparison.OrdinalIgnoreCase) = True Then
                    bomLevel = PartsListLevelEnum.kPartsOnly
                Else
                    bomLevel = PartsListLevelEnum.kStructured
                End If
            End If

            ' 7. Insertar la Lista de Piezas en la hoja
            Dim oPartsList As PartsList = Nothing
            Try
                If esPieza Then
                    oPartsList = targetSheet.PartsLists.Add(oView, oPoint)
                Else
                    oPartsList = targetSheet.PartsLists.Add(oView, oPoint, bomLevel)
                End If
            Catch ex As Exception
                Log($"Fallo al llamar PartsLists.Add con nivel específico: {ex.Message}. Intentando firma base...")
                Try
                    oPartsList = targetSheet.PartsLists.Add(oView, oPoint)
                Catch ex2 As Exception
                    Log($"Fallo definitivo al crear PartsList: {ex2.Message}")
                    Return False
                End Try
            End Try

            If oPartsList Is Nothing Then
                Log("No se pudo instanciar la PartsList en la hoja.")
                Return False
            End If

            ' 8. Aplicar el Estilo de Lista de Piezas configurado en Opciones
            Dim estiloDeseado As String = If(esPieza, cfg.EstiloListaPiezasPieza, cfg.EstiloListaPiezasEnsamblaje)
            Dim estiloAplicado As PartsListStyle = BuscarEstiloListaPiezas(drawDoc, estiloDeseado, esPieza)

            If estiloAplicado IsNot Nothing Then
                Try
                    oPartsList.Style = estiloAplicado
                    Log($"Estilo aplicado exitosamente a PartsList: '{estiloAplicado.Name}'")
                Catch ex As Exception
                    Log($"No se pudo asignar el estilo '{estiloAplicado.Name}': {ex.Message}")
                End Try
            Else
                Log($"No se encontró el estilo '{estiloDeseado}' ni estilos similares en el documento.")
            End If

            ' 9. Alinear con precisión milimétrica la lista de piezas:
            ' La esquina inferior izquierda de la lista de piezas tiene que coincidir con la esquina superior izquierda del cajetin
            AlinearListaPiezas(oPartsList, targetSheet, cfg.UbicacionListaPiezas, tg, drawDoc)

            ' 10. Actualizar dibujo
            Try
                drawDoc.Update2(True)
            Catch
            End Try

            Log($"Lista de piezas insertada con éxito en hoja '{targetSheet.Name}' para {If(esPieza, "Pieza", "Ensamblaje")}.")
            Return True

        Catch ex As Exception
            Log($"Excepción general en AutoPartsListManager.Ejecutar: {ex.Message}")
            Return False
        Finally
            _isProcessing = False
        End Try
    End Function

    ''' <summary>
    ''' Método estático para ser invocado desde eventos con un breve retardo (debounce)
    ''' para permitir que Inventor finalice la transacción de colocación de la vista.
    ''' </summary>
    Public Shared Sub NotificarComandoFinalizado(Optional drawDoc As DrawingDocument = Nothing)
        If _isProcessing Then Return

        ' Usar Timer en el hilo UI de Windows Forms para dar 250ms a Inventor de finalizar el dibujo
        Dim delayTimer As New System.Windows.Forms.Timer()
        delayTimer.Interval = 250
        AddHandler delayTimer.Tick, Sub(sender As Object, ev As EventArgs)
                                        delayTimer.Stop()
                                        delayTimer.Dispose()
                                        Try
                                            Dim mgr As New AutoPartsListManager(_invApp)
                                            mgr.Ejecutar()
                                        Catch ex As Exception
                                            Log($"Error en delayTimer de AutoPartsListManager: {ex.Message}")
                                        End Try
                                    End Sub
        delayTimer.Start()
    End Sub

    ''' <summary>
    ''' Determina con máxima precisión si el modelo referenciado en una vista es Pieza (.ipt) o Ensamblaje (.iam).
    ''' </summary>
    Public Shared Function DeterminarTipoComponente(view As DrawingView) As DocumentTypeEnum
        If view Is Nothing Then Return DocumentTypeEnum.kUnknownDocumentObject

        ' 1. A través de ReferencedDocumentDescriptor
        Try
            If view.ReferencedDocumentDescriptor IsNot Nothing Then
                Dim docType = view.ReferencedDocumentDescriptor.ReferencedDocumentType
                If docType = DocumentTypeEnum.kPartDocumentObject OrElse docType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Return docType
                End If

                If view.ReferencedDocumentDescriptor.ReferencedDocument IsNot Nothing Then
                    Dim rDoc = view.ReferencedDocumentDescriptor.ReferencedDocument
                    If rDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject OrElse rDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Return rDoc.DocumentType
                    End If
                End If

                Dim fullDocName As String = view.ReferencedDocumentDescriptor.FullDocumentName
                If Not String.IsNullOrEmpty(fullDocName) Then
                    If fullDocName.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then Return DocumentTypeEnum.kPartDocumentObject
                    If fullDocName.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then Return DocumentTypeEnum.kAssemblyDocumentObject
                End If
            End If
        Catch
        End Try

        ' 2. A través de ReferencedFileDescriptor
        Try
            If view.ReferencedFileDescriptor IsNot Nothing Then
                Dim fName As String = view.ReferencedFileDescriptor.FullFileName
                If Not String.IsNullOrEmpty(fName) Then
                    If fName.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then Return DocumentTypeEnum.kPartDocumentObject
                    If fName.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then Return DocumentTypeEnum.kAssemblyDocumentObject
                End If
            End If
        Catch
        End Try

        ' 3. A través de ReferencedFile
        Try
            If view.ReferencedFile IsNot Nothing Then
                Dim fName As String = view.ReferencedFile.FullFileName
                If Not String.IsNullOrEmpty(fName) Then
                    If fName.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then Return DocumentTypeEnum.kPartDocumentObject
                    If fName.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then Return DocumentTypeEnum.kAssemblyDocumentObject
                End If
            End If
        Catch
        End Try

        Return DocumentTypeEnum.kUnknownDocumentObject
    End Function

    ''' <summary>
    ''' Calcula el punto 2D de inserción de la tabla en base a la configuración y los límites de la hoja/cajetín.
    ''' </summary>
    Private Function CalcularPuntoColocacion(sheet As Sheet, ubicacion As String, tg As TransientGeometry) As Point2d
        Dim ubic As String = If(Not String.IsNullOrEmpty(ubicacion), ubicacion.Trim(), "SobreRotulo")

        Select Case ubic
            Case "SobreRotulo"
                ' Sobre el rótulo / cajetín:
                ' La esquina inferior izquierda de la lista de piezas debe coincidir con la esquina superior izquierda del cajetín.
                ' Esquina superior izquierda del TitleBlock:
                ' X = MinPoint.X (borde izquierdo del cajetín)
                ' Y = MaxPoint.Y (borde superior del cajetín)
                If sheet.TitleBlock IsNot Nothing Then
                    Return tg.CreatePoint2d(sheet.TitleBlock.RangeBox.MinPoint.X, sheet.TitleBlock.RangeBox.MaxPoint.Y)
                ElseIf sheet.Border IsNot Nothing Then
                    Return tg.CreatePoint2d(sheet.Border.RangeBox.MinPoint.X, sheet.Border.RangeBox.MinPoint.Y)
                Else
                    Return tg.CreatePoint2d(sheet.Width, sheet.Height)
                End If

            Case "SuperiorIzquierda"
                If sheet.Border IsNot Nothing Then
                    Return tg.CreatePoint2d(sheet.Border.RangeBox.MinPoint.X, sheet.Border.RangeBox.MaxPoint.Y)
                Else
                    Return tg.CreatePoint2d(0, sheet.Height)
                End If

            Case "InferiorIzquierda"
                If sheet.Border IsNot Nothing Then
                    Return tg.CreatePoint2d(sheet.Border.RangeBox.MinPoint.X, sheet.Border.RangeBox.MinPoint.Y)
                Else
                    Return tg.CreatePoint2d(0, 0)
                End If

            Case Else
                ' "SuperiorDerecha" (Predeterminado)
                If sheet.Border IsNot Nothing Then
                    Return sheet.Border.RangeBox.MaxPoint
                Else
                    Return tg.CreatePoint2d(sheet.Width, sheet.Height)
                End If
        End Select
    End Function

    ''' <summary>
    ''' Ajusta con precisión milimétrica la posición de la Lista de Piezas para garantizar que
    ''' la esquina inferior izquierda de la tabla coincida exactamente con la esquina superior izquierda del rótulo/cajetín.
    ''' </summary>
    Private Sub AlinearListaPiezas(oPartsList As PartsList, sheet As Sheet, ubicacion As String, tg As TransientGeometry, drawDoc As DrawingDocument)
        Try
            If oPartsList Is Nothing OrElse sheet Is Nothing Then Return

            Dim ubic As String = If(Not String.IsNullOrEmpty(ubicacion), ubicacion.Trim(), "SobreRotulo")
            If Not ubic.Equals("SobreRotulo", StringComparison.OrdinalIgnoreCase) Then Return

            If sheet.TitleBlock Is Nothing Then Return

            ' Actualizar documento para que Inventor calcule las dimensiones reales del RangeBox con el estilo aplicado
            Try
                If drawDoc IsNot Nothing Then drawDoc.Update2(True)
            Catch
            End Try

            ' Esquina superior izquierda del cajetín:
            ' X = MinPoint.X (borde izquierdo), Y = MaxPoint.Y (borde superior)
            Dim tbLeftX As Double = sheet.TitleBlock.RangeBox.MinPoint.X
            Dim tbTopY As Double = sheet.TitleBlock.RangeBox.MaxPoint.Y

            ' Esquina inferior izquierda actual de la lista de piezas:
            ' X = MinPoint.X (borde izquierdo), Y = MinPoint.Y (borde inferior)
            Dim plLeftX As Double = oPartsList.RangeBox.MinPoint.X
            Dim plBottomY As Double = oPartsList.RangeBox.MinPoint.Y

            Dim deltaX As Double = tbLeftX - plLeftX
            Dim deltaY As Double = tbTopY - plBottomY

            Log($"Alineación PartsList vs Cajetín: Cajetín Top-Left=({tbLeftX:F3}, {tbTopY:F3}) | PartsList Bottom-Left=({plLeftX:F3}, {plBottomY:F3}) | Delta=({deltaX:F3}, {deltaY:F3})")

            If Math.Abs(deltaX) > 0.001 OrElse Math.Abs(deltaY) > 0.001 Then
                Dim currentPos As Point2d = oPartsList.Position
                oPartsList.Position = tg.CreatePoint2d(currentPos.X + deltaX, currentPos.Y + deltaY)
                Log($"Posición corregida de PartsList. Nueva esquina inferior izquierda en ({oPartsList.RangeBox.MinPoint.X:F3}, {oPartsList.RangeBox.MinPoint.Y:F3})")
            End If
        Catch ex As Exception
            Log($"Error al alinear lista de piezas sobre el cajetín: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Localiza el estilo de lista de piezas en el StylesManager del plano con múltiples estrategias de fallback.
    ''' </summary>
    Private Function BuscarEstiloListaPiezas(drawDoc As DrawingDocument, nombreEstiloBuscado As String, esPieza As Boolean) As PartsListStyle
        If drawDoc Is Nothing OrElse drawDoc.StylesManager Is Nothing Then Return Nothing
        Dim styles = drawDoc.StylesManager.PartsListStyles

        ' Estrategia 1: Coincidencia exacta de nombre (ignorando mayúsculas/minúsculas)
        If Not String.IsNullOrWhiteSpace(nombreEstiloBuscado) Then
            For Each s As PartsListStyle In styles
                If s.Name.Equals(nombreEstiloBuscado.Trim(), StringComparison.OrdinalIgnoreCase) Then
                    Return s
                End If
            Next
        End If

        ' Estrategia 2: Coincidencia parcial con el nombre buscado
        If Not String.IsNullOrWhiteSpace(nombreEstiloBuscado) Then
            Dim cleanSearch = nombreEstiloBuscado.Trim().ToUpper()
            For Each s As PartsListStyle In styles
                If s.Name.ToUpper().Contains(cleanSearch) Then
                    Return s
                End If
            Next
        End If

        ' Estrategia 3: Fallback semántico para Piezas ("PIEZA", "PIEZAS", "LDM")
        If esPieza Then
            For Each s As PartsListStyle In styles
                Dim upper = s.Name.ToUpper()
                If upper.Contains("COPIA DE PIEZAS") OrElse
                   upper.Contains("PIEZAS") OrElse
                   upper.Contains("PIEZA") OrElse
                   upper = "LDM" Then
                    Return s
                End If
            Next
        Else
            ' Fallback semántico para Ensamblajes ("ENSAMBLAJE", "LDMARM", "LDM")
            For Each s As PartsListStyle In styles
                Dim upper = s.Name.ToUpper()
                If upper.Contains("COPIA DE ENSAMBLAJE") OrElse
                   upper.Contains("ENSAMBLAJE") OrElse
                   upper.Contains("LDMARM") OrElse
                   upper.Contains("LDM") Then
                    Return s
                End If
            Next
        End If

        ' Estrategia 4: Estilo predeterminado de la norma activa
        Try
            If drawDoc.StylesManager.ActiveStandardStyle IsNot Nothing Then
                ' Retorna Nothing para que conserve el estilo activo de la norma si no hay otro
            End If
        Catch
        End Try

        Return Nothing
    End Function

    ''' <summary>
    ''' Registra mensajes en el log de auditoría del complemento.
    ''' </summary>
    Private Shared Sub Log(mensaje As String)
        Try
            Dim dir As String = IOPath.GetDirectoryName(LOG_PATH)
            If Not System.IO.Directory.Exists(dir) Then System.IO.Directory.CreateDirectory(dir)
            Using sw As New System.IO.StreamWriter(LOG_PATH, True)
                sw.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | [AutoPartsList] {mensaje}")
            End Using
        Catch
        End Try
        System.Diagnostics.Debug.WriteLine($"[AutoPartsList] {mensaje}")
    End Sub

End Class
