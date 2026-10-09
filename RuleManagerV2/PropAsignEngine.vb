Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Windows.Forms
Imports Inventor

Imports IOPath = System.IO.Path
Imports IOFile = System.IO.File

''' <summary>
''' Motor nativo VB.NET para la asignación y cálculo de propiedades iProperties (MC, CYP, MEC, ARM, LARGO, ANCHO, ESPESOR).
''' Reemplaza de forma nativa la regla iLogic "PropAsign3.iLogicVb" e integra directamente el diálogo "SobremedidaLargoDialog.vb".
''' </summary>
Public Class PropAsignEngine

    Private ReadOnly _invApp As Inventor.Application

    Public Sub New(invApp As Inventor.Application)
        _invApp = invApp
    End Sub

    ''' <summary>
    ''' Ejecuta el proceso completo de análisis de subensamblajes y piezas,
    ''' asignación de propiedades y despliegue del diálogo de sobremedida (+3mm).
    ''' </summary>
    Public Function Ejecutar() As Boolean
        Try
            Dim doc As Document = _invApp.ActiveDocument
            If doc Is Nothing Then
                MessageBox.Show("No hay ningún documento activo en Inventor.", "Asignar Propiedades", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            If doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("El documento activo debe ser un ensamblaje (.iam) para asignar propiedades.", "Asignar Propiedades", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            Dim asmDoc As AssemblyDocument = DirectCast(doc, AssemblyDocument)

            ' Obtener metadatos del proyecto desde el ensamblaje principal (equivalente a PropAsignaux)
            Dim standardProps As PropertySet = Nothing
            Dim descripcion As String = ""
            Dim cliente As String = ""
            Dim proyectista As String = ""

            Try
                standardProps = asmDoc.PropertySets.Item("Design Tracking Properties")
                If standardProps IsNot Nothing Then
                    Try : descripcion = CStr(standardProps.Item("Project").Value) : Catch : End Try
                    Try : cliente = CStr(standardProps.Item("Vendor").Value) : Catch : End Try
                    Try : proyectista = CStr(standardProps.Item("Designer").Value) : Catch : End Try
                End If
            Catch
            End Try

            Dim listaMECItems As New List(Of ElementoMecItem)()
            Dim dictConsultasMEC As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim subAssemblyNames As New List(Of String)()

            ' Barra de progreso nativa de Inventor
            Dim oProgressBar As Inventor.ProgressBar = Nothing
            Try
                Dim totalOccs As Integer = asmDoc.ComponentDefinition.Occurrences.Count
                oProgressBar = _invApp.CreateProgressBar(True, totalOccs, "Asignando propiedades...")
            Catch
            End Try

            ' Analizar ensamblaje y subensamblajes recursivamente
            AnalyzeSubAssemblies(asmDoc.ComponentDefinition, subAssemblyNames, descripcion, cliente, listaMECItems, dictConsultasMEC, oProgressBar)

            Try
                If oProgressBar IsNot Nothing Then
                    oProgressBar.Close()
                End If
            Catch
            End Try

            ' Si se recolectaron piezas MEC, mostrar el diálogo interactivo de forma no modal (modeless)
            ' para permitir interactuar libremente con Inventor y con los documentos que se abran.
            If listaMECItems.Count > 0 Then
                Dim hwnd As IntPtr = IntPtr.Zero
                Try
                    hwnd = New IntPtr(_invApp.MainFrameHWND)
                Catch
                End Try

                Dim dlg As New SobremedidaLargoDialog(_invApp, listaMECItems)
                If hwnd <> IntPtr.Zero Then
                    dlg.Show(New WindowWrapper(hwnd))
                Else
                    dlg.Show()
                End If
            Else
                MessageBox.Show("Se han asignado las propiedades correctamente." & vbCrLf & vbCrLf &
                                "• Subensamblajes analizados: " & subAssemblyNames.Count & vbCrLf &
                                "• No se detectaron piezas MEC para sobremedida.",
                                "Asignar Propiedades", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("Error durante la asignación de propiedades: " & ex.Message & vbCrLf & ex.StackTrace,
                            "Error - Asignar Propiedades", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Recorre recursivamente todas las ocurrencias del ensamblaje o subensamblaje.
    ''' </summary>
    Private Sub AnalyzeSubAssemblies(compDef As ComponentDefinition,
                                    ByRef subAssemblyNames As List(Of String),
                                    descripcion As String,
                                    cliente As String,
                                    listaMECItems As List(Of ElementoMecItem),
                                    dictConsultasMEC As HashSet(Of String),
                                    oProgressBar As Inventor.ProgressBar)

        Dim oProgressStep As Integer = 0
        Dim totalSteps As Integer = compDef.Occurrences.Count

        For Each oOcc As ComponentOccurrence In compDef.Occurrences
            If oOcc.Suppressed Then Continue For
            If Not oOcc.Visible Then Continue For
            If oOcc.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then Continue For

            ' Omitir miembros del Centro de Contenido
            If oOcc.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                Try
                    If oOcc.Definition.IsContentMember Then
                        Continue For
                    End If
                Catch
                End Try
            End If

            ' Actualizar progreso
            oProgressStep += 1
            Try
                If oProgressBar IsNot Nothing Then
                    oProgressBar.Message = "Procesando (" & oProgressStep & " de " & totalSteps & "): " & oOcc.Name
                    oProgressBar.UpdateProgress()
                End If
            Catch
            End Try

            If oOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                subAssemblyNames.Add(oOcc.Name)

                Dim subAssemblyCompDef As AssemblyComponentDefinition = Nothing
                Try
                    subAssemblyCompDef = DirectCast(oOcc.Definition, AssemblyComponentDefinition)
                Catch
                End Try

                ' Procesar las propiedades del subensamblaje como elemento
                ProcesarOcurrencia(oOcc, descripcion, cliente, listaMECItems, dictConsultasMEC)

                ' Recursión en los hijos del subensamblaje
                If subAssemblyCompDef IsNot Nothing Then
                    AnalyzeSubAssemblies(subAssemblyCompDef, subAssemblyNames, descripcion, cliente, listaMECItems, dictConsultasMEC, oProgressBar)
                End If
            Else
                ' Procesar pieza individual
                ProcesarOcurrencia(oOcc, descripcion, cliente, listaMECItems, dictConsultasMEC)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Configura las propiedades personalizadas base (MC, CYP, MEC, ARM) y delega la asignación según tipo.
    ''' </summary>
    Private Sub ProcesarOcurrencia(elemento As ComponentOccurrence,
                                   descripcion As String,
                                   cliente As String,
                                   listaMECItems As List(Of ElementoMecItem),
                                   dictConsultasMEC As HashSet(Of String))
        Try
            Dim customProps As PropertySet = Nothing
            Try
                customProps = elemento.Definition.Document.PropertySets.Item("Inventor User Defined Properties")
            Catch
                customProps = elemento.Definition.Document.PropertySets.Add("Inventor User Defined Properties")
            End Try

            Dim propMC As Inventor.Property = GetOrCreateProperty(customProps, "MC", elemento)
            Dim propCYP As Inventor.Property = GetOrCreateProperty(customProps, "CYP", elemento)
            Dim propMEC As Inventor.Property = GetOrCreateProperty(customProps, "MEC", elemento)
            Dim propARM As Inventor.Property = GetOrCreateProperty(customProps, "ARM", elemento)

            Asignar(propMC, propCYP, propMEC, propARM, elemento, customProps, descripcion, cliente, listaMECItems, dictConsultasMEC)
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Evalúa si el elemento es comercial (COM), chapa (MC/CYP), ensamble (ARM) o mecanizado (MEC),
    ''' calculando dimensiones y metadatos correspondientes.
    ''' </summary>
    Private Sub Asignar(propMC As Inventor.Property,
                        propCYP As Inventor.Property,
                        propMEC As Inventor.Property,
                        propARM As Inventor.Property,
                        elemento As ComponentOccurrence,
                        customProps As PropertySet,
                        descripcion As String,
                        cliente As String,
                        listaMECItems As List(Of ElementoMecItem),
                        dictConsultasMEC As HashSet(Of String))

        Dim standardProperties As PropertySet = Nothing
        Try
            standardProperties = elemento.Definition.Document.PropertySets.Item("Design Tracking Properties")
        Catch
        End Try

        ' 1. Verificar si la propiedad COM es "✓"
        Try
            Dim propCOM As Inventor.Property = customProps.Item("COM")
            If propCOM IsNot Nothing AndAlso CStr(propCOM.Value) = "✓" Then
                TryDeleteProperty(customProps, "MC")
                TryDeleteProperty(customProps, "CYP")
                TryDeleteProperty(customProps, "ARM")
                TryDeleteProperty(customProps, "MEC")

                Dim totalLengthstr As String = CalculateTotalLength(elemento)
                Dim propLARGO As Inventor.Property = GetOrCreateProperty(customProps, "LARGO", elemento)
                propLARGO.Value = totalLengthstr

                Dim totalWidthstr As String = CalculateTotalWidth(elemento)
                Dim propANCHO As Inventor.Property = GetOrCreateProperty(customProps, "ANCHO", elemento)
                propANCHO.Value = totalWidthstr

                Exit Sub
            End If
        Catch
        End Try

        ' 2. Asignación según el tipo de componente
        Dim defType As Integer = CInt(elemento.Definition.Type)

        Select Case defType
            Case 100663808 ' CÓDIGO IAM (Subensamblaje)
                If propARM IsNot Nothing Then propARM.Value = "✓"

                Dim totalLengthstr As String = CalculateTotalLength(elemento)
                Dim propLARGO As Inventor.Property = GetOrCreateProperty(customProps, "LARGO", elemento)
                propLARGO.Value = totalLengthstr

                Dim totalWidthstr As String = CalculateTotalWidth(elemento)
                Dim propANCHO As Inventor.Property = GetOrCreateProperty(customProps, "ANCHO", elemento)
                propANCHO.Value = totalWidthstr

                If standardProperties IsNot Nothing Then
                    SetPropertyValueSafe(standardProperties, "Project", descripcion)
                    SetPropertyValueSafe(standardProperties, "Vendor", cliente)
                End If

            Case 150995200 ' CÓDIGO PIEZA CHAPA (Sheet Metal)
                If propMC IsNot Nothing Then propMC.Value = "✓"

                Dim oSheetMetalCompDef As SheetMetalComponentDefinition = TryCast(elemento.Definition, SheetMetalComponentDefinition)
                If oSheetMetalCompDef IsNot Nothing Then
                    If Not oSheetMetalCompDef.HasFlatPattern Then
                        Try
                            oSheetMetalCompDef.Unfold()
                        Catch
                        End Try
                    End If

                    Dim propMAT_PLANO As Inventor.Property = GetOrCreateProperty(customProps, "MATERIAL_PLANO", elemento)
                    If String.IsNullOrWhiteSpace(CStr(propMAT_PLANO.Value)) AndAlso standardProperties IsNot Nothing Then
                        Try
                            propMAT_PLANO.Value = standardProperties.Item("Material").Value
                        Catch
                        End Try
                    End If

                    Dim propLARGO As Inventor.Property = GetOrCreateProperty(customProps, "LARGO", elemento)
                    propLARGO.Value = "=<Flat Pattern Length>"
                    Try
                        Dim doubleVal As Double = CDbl(propLARGO.Value) * 10.0
                        doubleVal = Math.Round(doubleVal, 0)
                        propLARGO.Value = doubleVal.ToString() & " mm"
                    Catch
                    End Try

                    Dim propANCHO As Inventor.Property = GetOrCreateProperty(customProps, "ANCHO", elemento)
                    propANCHO.Value = "=<Flat Pattern Width>"
                    Try
                        Dim doubleVal As Double = CDbl(propANCHO.Value) * 10.0
                        doubleVal = Math.Round(doubleVal, 0)
                        propANCHO.Value = doubleVal.ToString() & " mm"
                    Catch
                    End Try

                    If oSheetMetalCompDef.FlatPattern IsNot Nothing AndAlso oSheetMetalCompDef.FlatPattern.FlatBendResults IsNot Nothing Then
                        Dim numPatrones As Integer = oSheetMetalCompDef.FlatPattern.FlatBendResults.Count
                        Dim numPliegues As Integer = numPatrones \ 2
                        If numPliegues > 0 AndAlso propCYP IsNot Nothing Then
                            propCYP.Value = "✓"
                        End If
                    End If

                    If standardProperties IsNot Nothing Then
                        SetPropertyValueSafe(standardProperties, "Project", descripcion)
                        SetPropertyValueSafe(standardProperties, "Vendor", cliente)
                    End If

                    Dim propESPESOR As Inventor.Property = GetOrCreateProperty(customProps, "ESPESOR", elemento)
                    Try
                        propESPESOR.Value = oSheetMetalCompDef.Thickness.Value * 10.0
                    Catch
                    End Try
                End If

            Case 83886592 ' CÓDIGO PIEZA NORMALIZADA (Standard Part)
                If propMEC IsNot Nothing Then propMEC.Value = "✓"

                Try
                    Dim propPerf As Inventor.Property = customProps.Item("PERFILERIA")
                    If propPerf IsNot Nothing AndAlso CStr(propPerf.Value) = "✓" Then
                        If propMEC IsNot Nothing Then propMEC.Value = ""
                    End If
                Catch
                End Try

                Dim propMAT_PLANO As Inventor.Property = GetOrCreateProperty(customProps, "MATERIAL_PLANO", elemento)
                If String.IsNullOrWhiteSpace(CStr(propMAT_PLANO.Value)) AndAlso standardProperties IsNot Nothing Then
                    Try
                        propMAT_PLANO.Value = standardProperties.Item("Material").Value
                    Catch
                    End Try
                End If

                Dim totalLengthstr As String = CalculateTotalLength(elemento)
                Dim propLARGO As Inventor.Property = GetOrCreateProperty(customProps, "LARGO", elemento)
                propLARGO.Value = totalLengthstr

                Dim totalWidthstr As String = CalculateTotalWidth(elemento)
                Dim propANCHO As Inventor.Property = GetOrCreateProperty(customProps, "ANCHO", elemento)
                propANCHO.Value = totalWidthstr

                If standardProperties IsNot Nothing Then
                    SetPropertyValueSafe(standardProperties, "Project", descripcion)
                    SetPropertyValueSafe(standardProperties, "Vendor", cliente)
                End If

                ' Si la propiedad MEC tiene el valor "✓", recolectar pieza para la revisión de sobremedida (+3mm)
                If propMEC IsNot Nothing AndAlso CStr(propMEC.Value) = "✓" Then
                    RecolectarElementoMec(elemento, standardProperties, customProps, propMAT_PLANO, propLARGO, totalLengthstr, totalWidthstr, listaMECItems, dictConsultasMEC)
                End If
        End Select
    End Sub

    ''' <summary>
    ''' Registra un item único para el diálogo de sobremedida de largo (+3mm).
    ''' </summary>
    Private Sub RecolectarElementoMec(elemento As ComponentOccurrence,
                                       standardProperties As PropertySet,
                                       customProps As PropertySet,
                                       propMAT_PLANO As Inventor.Property,
                                       propLARGO As Inventor.Property,
                                       totalLengthstr As String,
                                       totalWidthstr As String,
                                       listaMECItems As List(Of ElementoMecItem),
                                       dictConsultasMEC As HashSet(Of String))

        Dim docPath As String = ""
        Try
            docPath = elemento.Definition.Document.FullFileName
        Catch
            docPath = elemento.Name
        End Try

        If dictConsultasMEC.Contains(docPath) Then Return
        dictConsultasMEC.Add(docPath)

        Dim sPartNumber As String = ""
        If standardProperties IsNot Nothing Then
            Try
                sPartNumber = CStr(standardProperties.Item("Part Number").Value)
            Catch
                sPartNumber = elemento.Name
            End Try
        Else
            sPartNumber = elemento.Name
        End If

        Dim sStockNumber As String = ""
        If standardProperties IsNot Nothing Then
            Try
                sStockNumber = CStr(standardProperties.Item("Stock Number").Value)
            Catch
            End Try
        End If
        If String.IsNullOrEmpty(sStockNumber) AndAlso customProps IsNot Nothing Then
            Try
                sStockNumber = CStr(customProps.Item("Stock Number").Value)
            Catch
            End Try
        End If

        Dim sMatPlano As String = ""
        If propMAT_PLANO IsNot Nothing Then
            Try
                sMatPlano = CStr(propMAT_PLANO.Value)
            Catch
            End Try
        End If

        Dim numLargo As Double = 0
        Dim strLimpio As String = totalLengthstr.Replace("mm", "").Trim()
        If Not Double.TryParse(strLimpio, numLargo) Then
            Try
                numLargo = CDbl(strLimpio)
            Catch
                numLargo = 0
            End Try
        End If

        Dim itemMec As New ElementoMecItem With {
            .Occurrence = elemento,
            .Document = elemento.Definition.Document,
            .DocumentPath = docPath,
            .FileName = IOPath.GetFileName(docPath),
            .PartNumber = sPartNumber,
            .StockNumber = sStockNumber,
            .MaterialPlano = sMatPlano,
            .LargoOriginalStr = totalLengthstr,
            .LargoOriginalNum = numLargo,
            .AnchoStr = totalWidthstr,
            .Tipo = "MEC",
            .Sumar3mm = True,
            .PropLargo = propLARGO
        }

        listaMECItems.Add(itemMec)
    End Sub

    ''' <summary>
    ''' Calcula la longitud mayor de la caja delimitadora precisa (PreciseRangeBox) en mm redondeados.
    ''' </summary>
    Private Function CalculateTotalLength(elemento As ComponentOccurrence) As String
        Try
            Dim cajaDelimitadora As Box = elemento.Definition.Document.ComponentDefinition.PreciseRangeBox
            Dim Xlength As Double = (cajaDelimitadora.MaxPoint.X - cajaDelimitadora.MinPoint.X) * 10.0
            Dim Ylength As Double = (cajaDelimitadora.MaxPoint.Y - cajaDelimitadora.MinPoint.Y) * 10.0
            Dim Zlength As Double = (cajaDelimitadora.MaxPoint.Z - cajaDelimitadora.MinPoint.Z) * 10.0
            Dim totalLength As Double

            If Xlength = Ylength Then
                Xlength = 0
                Ylength = 0
            ElseIf Xlength = Zlength Then
                Xlength = 0
                Zlength = 0
            ElseIf Ylength = Zlength Then
                Ylength = 0
                Zlength = 0
            End If

            If Xlength > Ylength AndAlso Xlength > Zlength Then
                totalLength = Xlength
            ElseIf Ylength > Xlength AndAlso Ylength > Zlength Then
                totalLength = Ylength
            Else
                totalLength = Zlength
            End If

            totalLength = Math.Round(totalLength)
            Return totalLength.ToString() & " mm"
        Catch
            Return "0 mm"
        End Try
    End Function

    ''' <summary>
    ''' Calcula el ancho (segunda dimensión mayor) de la caja delimitadora precisa en mm redondeados.
    ''' </summary>
    Private Function CalculateTotalWidth(elemento As ComponentOccurrence) As String
        Try
            Dim cajaDelimitadora As Box = elemento.Definition.Document.ComponentDefinition.PreciseRangeBox
            Dim XWidth As Double = (cajaDelimitadora.MaxPoint.X - cajaDelimitadora.MinPoint.X) * 10.0
            Dim YWidth As Double = (cajaDelimitadora.MaxPoint.Y - cajaDelimitadora.MinPoint.Y) * 10.0
            Dim ZWidth As Double = (cajaDelimitadora.MaxPoint.Z - cajaDelimitadora.MinPoint.Z) * 10.0
            Dim totalWidth As Double

            If (XWidth >= YWidth AndAlso XWidth <= ZWidth) OrElse (XWidth <= YWidth AndAlso XWidth >= ZWidth) Then
                totalWidth = XWidth
            ElseIf (YWidth >= XWidth AndAlso YWidth <= ZWidth) OrElse (YWidth <= XWidth AndAlso YWidth >= ZWidth) Then
                totalWidth = YWidth
            Else
                totalWidth = ZWidth
            End If

            totalWidth = Math.Round(totalWidth)
            Return totalWidth.ToString() & " mm"
        Catch
            Return "0 mm"
        End Try
    End Function

    ''' <summary>
    ''' Obtiene una propiedad existente en el PropertySet o la crea si no existe.
    ''' </summary>
    Private Function GetOrCreateProperty(props As PropertySet, propName As String, Optional oOcc As ComponentOccurrence = Nothing) As Inventor.Property
        If props Is Nothing Then Return Nothing
        For Each prop As Inventor.Property In props
            If String.Equals(prop.Name, propName, StringComparison.OrdinalIgnoreCase) Then
                Return prop
            End If
        Next
        Try
            Return props.Add(" ", propName)
        Catch ex As Exception
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Elimina una propiedad si existe de forma segura.
    ''' </summary>
    Private Sub TryDeleteProperty(props As PropertySet, propName As String)
        If props Is Nothing Then Return
        Try
            Dim p As Inventor.Property = props.Item(propName)
            If p IsNot Nothing Then p.Delete()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Asigna un valor a una propiedad de forma segura evitando excepciones.
    ''' </summary>
    Private Sub SetPropertyValueSafe(props As PropertySet, propName As String, val As Object)
        If props Is Nothing Then Return
        Try
            props.Item(propName).Value = val
        Catch
        End Try
    End Sub

End Class
