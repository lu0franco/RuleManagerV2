Imports System.Windows.Forms
Imports Inventor
Imports IODirectory = System.IO.Directory
Imports IOFile = System.IO.File
Imports IOPath = System.IO.Path

Public Class Rules

    Private Shared Function GetILogic(
        invApp As Inventor.Application) As Object

        Return invApp.ApplicationAddIns.ItemById(
            "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}"
        ).Automation

    End Function


    ' REEMPLAZAR el método actual:
    Public Shared Sub EjecutarSComAsign(invApp As Inventor.Application)
        Try
            ' Opción B: Con visualizador (recomendado)
            Dim dialog As New ComercialAssignmentDialog(invApp)
            dialog.ShowDialog()

        Catch ex As Exception
            MsgBox("Error en Asignar Comerciales: " & ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Public Shared Sub EjecutarPropAsign(
        invApp As Inventor.Application)

        Try

            Dim iLogicAuto As Object =
                GetILogic(invApp)

            iLogicAuto.RunExternalRule(
                invApp.ActiveDocument,
                "D:\Bibliotecas\Ilogic\Reglas\IAM\PropAsign3.iLogicVb"
            )

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub


    Public Shared Sub EjecutarSim(
        invApp As Inventor.Application)

        Try

            Dim iLogicAuto As Object =
                GetILogic(invApp)

            iLogicAuto.RunExternalRule(
                invApp.ActiveDocument,
                "D:\Bibliotecas\Ilogic\Reglas\IAM\SEARCHSIMETRIAS.iLogicVb"
            )

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Shared Sub EjecutarPartsCount(
        invApp As Inventor.Application)

        Try

            Dim iLogicAuto As Object =
                GetILogic(invApp)

            iLogicAuto.RunExternalRule(
                invApp.ActiveDocument,
                "D:\Bibliotecas\Ilogic\Reglas\IAM\Contador de elementosV2.iLogicVb"
            )

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Shared Sub EjecutarRevSim(
        invApp As Inventor.Application)

        Try

            Dim iLogicAuto As Object =
                GetILogic(invApp)

            iLogicAuto.RunExternalRule(
                invApp.ActiveDocument,
                "D:\Bibliotecas\Ilogic\Reglas\IAM\RevisionTEXTO_CANTIDADES.iLogicVb"
            )

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Shared Sub EjecutarSelBase(
        invApp As Inventor.Application)

        Try

            Dim iLogicAuto As Object =
                GetILogic(invApp)

            iLogicAuto.RunExternalRule(
                invApp.ActiveDocument,
                "D:\Bibliotecas\Ilogic\Reglas\IPT\Seleccion baseV2.iLogicVb"
            )

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub
    Public Shared Sub EjecutarPropAsignPart(
       invApp As Inventor.Application)

        Try

            Dim iLogicAuto As Object =
                GetILogic(invApp)

            iLogicAuto.RunExternalRule(
                invApp.ActiveDocument,
                "D:\Bibliotecas\Ilogic\Reglas\IPT\PropAsignPiezas.iLogicVb"
            )

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub
    Public Shared Sub EjecutarVincularDatos(invApp As Inventor.Application)
        Try
            If invApp.ActiveDocument Is Nothing Then
                MessageBox.Show("No hay ningún documento activo en Inventor.", "Vincular Datos", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            If invApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                MessageBox.Show("Esta función solo se puede ejecutar en un documento de dibujo (.idw / .dwg).", "Vincular Datos", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            Dim drawDoc As DrawingDocument = CType(invApp.ActiveDocument, DrawingDocument)

            ' Instanciar ventana de apertura de archivos de Windows (OpenFileDialog)
            Dim selectedFilePath As String = ""
            Using ofd As New OpenFileDialog()
                ofd.Title = "Seleccionar archivo (Pieza o Ensamblaje) para vincular N° de pieza"
                ofd.Filter = "Modelos de Inventor (*.iam;*.ipt)|*.iam;*.ipt|Ensamblajes de Inventor (*.iam)|*.iam|Piezas de Inventor (*.ipt)|*.ipt|Todos los archivos (*.*)|*.*"
                ofd.FilterIndex = 1
                ofd.CheckFileExists = True
                ofd.Multiselect = False
                ofd.RestoreDirectory = True

                ' Definir carpeta inicial según el dibujo activo o el proyecto de Inventor
                If Not String.IsNullOrEmpty(drawDoc.FullFileName) Then
                    Try
                        ofd.InitialDirectory = IOPath.GetDirectoryName(drawDoc.FullFileName)
                    Catch
                    End Try
                Else
                    Try
                        If invApp.DesignProjectManager IsNot Nothing AndAlso
                           invApp.DesignProjectManager.ActiveDesignProject IsNot Nothing Then
                            ofd.InitialDirectory = invApp.DesignProjectManager.ActiveDesignProject.WorkspacePath
                        End If
                    Catch
                    End Try
                End If

                Dim hwnd As IntPtr = IntPtr.Zero
                Try
                    hwnd = New IntPtr(invApp.MainFrameHWND)
                Catch
                End Try

                Dim res As DialogResult
                If hwnd <> IntPtr.Zero Then
                    res = ofd.ShowDialog(New WindowWrapper(hwnd))
                Else
                    res = ofd.ShowDialog()
                End If

                If res <> DialogResult.OK OrElse String.IsNullOrWhiteSpace(ofd.FileName) Then
                    Exit Sub
                End If

                selectedFilePath = ofd.FileName
            End Using

            If Not IOFile.Exists(selectedFilePath) Then
                MessageBox.Show("El archivo seleccionado no existe.", "Vincular Datos", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            ' Verificar si el archivo ya está abierto en Inventor
            Dim modelDoc As Document = Nothing
            Dim openedInvisibly As Boolean = False

            For Each d As Document In invApp.Documents
                If String.Equals(d.FullFileName, selectedFilePath, StringComparison.OrdinalIgnoreCase) Then
                    modelDoc = d
                    Exit For
                End If
            Next

            ' Si no está abierto en la sesión, abrirlo de forma invisible
            If modelDoc Is Nothing Then
                Try
                    modelDoc = invApp.Documents.Open(selectedFilePath, False)
                    openedInvisibly = True
                Catch ex As Exception
                    MessageBox.Show("No se pudo abrir el archivo en Inventor:" & vbCrLf & ex.Message, "Error al abrir archivo", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End Try
            End If

            ' Obtener el N° de pieza (Part Number) de las iProperties del modelo
            Dim modelPartNumber As String = GetPropertyFromDoc(modelDoc, "Part Number")

            ' Si abrimos el archivo temporalmente de forma invisible, cerrarlo para no dejarlo abierto en memoria
            If openedInvisibly AndAlso modelDoc IsNot Nothing Then
                Try
                    modelDoc.Close(True)
                Catch
                End Try
            End If

            ' Si el N° de pieza está vacío, advertir al usuario
            If String.IsNullOrWhiteSpace(modelPartNumber) Then
                Dim resp As DialogResult = MessageBox.Show(
                    "El archivo seleccionado '" & IOPath.GetFileName(selectedFilePath) & "' no tiene un N° de pieza definido (está vacío)." & vbCrLf & vbCrLf &
                    "¿Desea asignar un valor vacío al N° de pieza del dibujo?",
                    "N° de Pieza Vacío",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question)
                If resp <> DialogResult.Yes Then
                    Exit Sub
                End If
            End If

            ' Asignar el N° de pieza en las iProperties del dibujo activo
            SetPropertyInDoc(drawDoc, "Description", modelPartNumber)

            ' Actualizar el dibujo para reflejar el cambio en cajetín/rótulo
            Try
                drawDoc.Update2(True)
            Catch
                Try
                    drawDoc.Update()
                Catch
                End Try
            End Try

            MessageBox.Show(
                "N° de pieza vinculado con éxito al dibujo:" & vbCrLf & vbCrLf &
                "• Archivo de origen: " & IOPath.GetFileName(selectedFilePath) & vbCrLf &
                "• N° de pieza asignado: " & modelPartNumber,
                "Vincular Datos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error en Vincular Datos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Shared Sub EjecutarTraerCantidades(
   invApp As Inventor.Application)

        Try

            Dim iLogicAuto As Object =
                GetILogic(invApp)

            iLogicAuto.RunExternalRule(
                invApp.ActiveDocument,
                "D:\Bibliotecas\Ilogic\Reglas\IDW\Traer cantidadesV2.iLogicVb"
            )

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Shared Sub EjecutarAddCopias(
   invApp As Inventor.Application)

        Try

            Dim iLogicAuto As Object =
                GetILogic(invApp)

            iLogicAuto.RunExternalRule(
                invApp.ActiveDocument,
                "D:\Bibliotecas\Ilogic\Reglas\IDW\Añadir copias.iLogicVb"
            )

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub


    Public Shared Sub EjecutarDelCopias(
   invApp As Inventor.Application)

        Try

            Dim iLogicAuto As Object =
                GetILogic(invApp)

            iLogicAuto.RunExternalRule(
                invApp.ActiveDocument,
                "D:\Bibliotecas\Ilogic\Reglas\IDW\borrar copias.iLogicVb"
            )

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub


    Public Shared Sub EjecutarCambiarReferencias(
   invApp As Inventor.Application)

        Try

            Dim iLogicAuto As Object =
                GetILogic(invApp)

            iLogicAuto.RunExternalRule(
                invApp.ActiveDocument,
                "D:\Bibliotecas\Ilogic\Reglas\IDW\Cambiar Referencias.iLogicVb"
            )

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub


    Public Shared Sub EjecutarOrdPlanos(
   invApp As Inventor.Application)

        Try

            Dim iLogicAuto As Object =
                GetILogic(invApp)

            iLogicAuto.RunExternalRule(
                invApp.ActiveDocument,
                "D:\Bibliotecas\Ilogic\Reglas\IDW\Ordenar PlanosV2.iLogicVb"
            )

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub


    Public Shared Sub EjecutarDelListas(
   invApp As Inventor.Application)

        Try

            Dim iLogicAuto As Object =
                GetILogic(invApp)

            iLogicAuto.RunExternalRule(
                invApp.ActiveDocument,
                "D:\Bibliotecas\Ilogic\Reglas\IDW\Limpiar listas.iLogicVb"
            )

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub
    Public Shared Sub EjecutarLDMMANAGER(
   invApp As Inventor.Application)

        Try

            Dim iLogicAuto As Object =
                GetILogic(invApp)

            iLogicAuto.RunExternalRule(
                invApp.ActiveDocument,
                "D:\Bibliotecas\Ilogic\Reglas\IDW\LDM_MANAGER_V5.iLogicVb"
            )

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Shared Sub EjecutarPrintLEGALes(
   invApp As Inventor.Application)

        Try

            Dim iLogicAuto As Object =
                GetILogic(invApp)

            iLogicAuto.RunExternalRule(
                invApp.ActiveDocument,
                "D:\Bibliotecas\Ilogic\Reglas\IDW\PrintLEGALes.iLogicVb"
            )

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Shared Sub EjecutarPrintB3(
   invApp As Inventor.Application)

        Try

            Dim iLogicAuto As Object =
                GetILogic(invApp)

            iLogicAuto.RunExternalRule(
                invApp.ActiveDocument,
                "D:\Bibliotecas\Ilogic\Reglas\IDW\Print B3.iLogicVb"
            )

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Shared Sub EjecutarPrintB3EXT(
   invApp As Inventor.Application)

        Try

            Dim iLogicAuto As Object =
                GetILogic(invApp)

            iLogicAuto.RunExternalRule(
                invApp.ActiveDocument,
                "D:\Bibliotecas\Ilogic\Reglas\IDW\Print B3 ext.iLogicVb"
            )

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Shared Sub EjecutarPrintB3EXTWIDE(
   invApp As Inventor.Application)

        Try

            Dim iLogicAuto As Object =
                GetILogic(invApp)

            iLogicAuto.RunExternalRule(
                invApp.ActiveDocument,
                "D:\Bibliotecas\Ilogic\Reglas\IDW\Print B3 ext Wide.iLogicVb"
            )

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    ' ============================================
    ' ORGANIZAR PROYECTO → StartupTasks
    ' ============================================
    Public Shared Sub EjecutarOrganizarProyecto(
        invApp As Inventor.Application)

        Try
            StartupTasks.OrganizarArchivosProyecto(invApp)
            MsgBox("Proyecto organizado correctamente.")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    ' ============================================
    ' RENOMBRAR PROYECTO - CON GUARDADO PREVIO
    ' ============================================
    Public Shared Sub EjecutarProjectRename(
    invApp As Inventor.Application)

        Try

            Dim asmDoc As Inventor.AssemblyDocument =
            CType(invApp.ActiveDocument, Inventor.AssemblyDocument)

            ' === PASO 1: GUARDAR DOCUMENTO ACTUAL ===
            Try
                asmDoc.Save2(False)
                System.Threading.Thread.Sleep(500)
            Catch ex As Exception
                ' Si no se puede guardar, continuar igual
            End Try

            ' === PASO 2: EXTRAER DATOS DEL DOCUMENTO ===
            Dim data As New ProjectData()
            data.Modelo = GetPropertyFromDoc(asmDoc, "Stock Number")
            data.Maquina = GetPropertyFromDoc(asmDoc, "Part Number")
            data.OT = GetPropertyFromDoc(asmDoc, "Project")
            data.Cliente = GetPropertyFromDoc(asmDoc, "Vendor")

            ' === PASO 3: MOSTRAR DIÁLOGO DE CONFIRMACIÓN/EDICIÓN ===
            Dim dialog As New ProjectRenameDialog(data)
            Dim result As DialogResult = dialog.ShowDialog()

            If result <> DialogResult.OK Then
                ' Usuario canceló, salir
                Exit Sub
            End If

            ' === PASO 4: SI HAY CAMBIOS, GUARDAR PROPIEDADES DE VUELTA ===
            Dim finalData As ProjectData = dialog.ResultData
            Dim cambios As Boolean = False

            If finalData.Modelo <> data.Modelo Then
                SetPropertyInDoc(asmDoc, "Stock Number", finalData.Modelo)
                cambios = True
            End If
            If finalData.Maquina <> data.Maquina Then
                SetPropertyInDoc(asmDoc, "Part Number", finalData.Maquina)
                cambios = True
            End If
            If finalData.OT <> data.OT Then
                SetPropertyInDoc(asmDoc, "Project", finalData.OT)
                cambios = True
            End If
            If finalData.Cliente <> data.Cliente Then
                SetPropertyInDoc(asmDoc, "Vendor", finalData.Cliente)
                cambios = True
            End If

            ' === PASO 5: RE-GUARDAR SI HUBO CAMBIOS ===
            If cambios Then
                Try
                    asmDoc.Save2(False)
                    System.Threading.Thread.Sleep(500)
                Catch ex As Exception
                    MsgBox("Error al guardar propiedades modificadas: " & ex.Message)
                    Exit Sub
                End Try
            End If

            ' === PASO 6: EJECUTAR WORKFLOW (ProjectRenamer NO SE TOCA) ===
            Dim workflow As New ProjectRenameWorkflow(
            invApp,
            asmDoc
        )

            Dim targetFolder As String = workflow.Execute()

            MsgBox("Proyecto renombrado correctamente en: " & targetFolder)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    ' === HELPERS PARA LEER/ESCRIBIR PROPIEDADES ===
    Private Shared Function GetPropertyFromDoc(doc As Document, propName As String) As String
        Try
            Return doc.PropertySets.Item("Design Tracking Properties").Item(propName).Value.ToString()
        Catch
            Try
                Return doc.PropertySets.Item("Inventor Summary Information").Item(propName).Value.ToString()
            Catch
                Try
                    Return doc.PropertySets.Item("Inventor User Defined Properties").Item(propName).Value.ToString()
                Catch
                    Return ""
                End Try
            End Try
        End Try
    End Function

    Private Shared Sub SetPropertyInDoc(doc As Document, propName As String, value As String)
        ' 1. Intentar en Design Tracking Properties (donde reside normalmente Part Number, Stock Number, etc.)
        Try
            Dim propsDT As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
            Dim propDT As Inventor.Property = propsDT.Item(propName)
            If propDT IsNot Nothing Then
                propDT.Value = value
                Return
            End If
        Catch
        End Try

        ' 2. Intentar en Inventor Summary Information (Title, Subject, Author, etc.)
        Try
            Dim propsSI As PropertySet = doc.PropertySets.Item("Inventor Summary Information")
            Dim propSI As Inventor.Property = propsSI.Item(propName)
            If propSI IsNot Nothing Then
                propSI.Value = value
                Return
            End If
        Catch
        End Try

        ' 3. Si no es estándar o no existe, escribir o agregar a Inventor User Defined Properties
        Try
            Dim propsUser As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
            Dim propUser As Inventor.Property = Nothing
            Try
                propUser = propsUser.Item(propName)
            Catch
                propUser = propsUser.Add(value, propName)
            End Try
            If propUser IsNot Nothing Then
                propUser.Value = value
            End If
        Catch ex As Exception
            Throw New Exception("No se pudo escribir la propiedad '" & propName & "': " & ex.Message)
        End Try
    End Sub

    ' ====================================================
    ' CORREGIDO: RENOMBRAR COMPONENTES DEL `.iam` ACTIVO
    ' ====================================================
    Public Shared Sub EjecutarRenombrarComponentes(invApp As Inventor.Application)

        Try
            ' 1. Validar que haya un documento abierto
            If invApp.ActiveDocument Is Nothing Then
                MsgBox("No hay ningún documento activo en Inventor.", MsgBoxStyle.Exclamation, "Atención")
                Exit Sub
            End If

            ' 2. Validar que el archivo activo sea realmente un ensamblaje (.iam)
            If invApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MsgBox("Esta función solo se puede ejecutar estando dentro de un Ensamblaje (.iam).", MsgBoxStyle.Exclamation, "Atención")
                Exit Sub
            End If

            ' 3. Castear el documento a AssemblyDocument de forma segura
            Dim oAsmDoc As AssemblyDocument = CType(invApp.ActiveDocument, AssemblyDocument)
            Dim sRutaActiva As String = oAsmDoc.FullFileName

            ' 4. Validar que el archivo esté guardado en disco (que tenga una ruta válida)
            If String.IsNullOrEmpty(sRutaActiva) OrElse Not IOFile.Exists(sRutaActiva) Then
                MsgBox("Por favor, guardá el ensamblaje antes de ejecutar el renombrador.", MsgBoxStyle.Critical, "Archivo no guardado")
                Exit Sub
            End If

            ' 5. Conseguir el prefijo dinámicamente usando tu helper leyendo la propiedad 'Project' u otra que uses
            Dim sPrefijo As String = GetPropertyFromDoc(oAsmDoc, "Project")
            If String.IsNullOrEmpty(sPrefijo) Then
                sPrefijo = "PROY" ' Prefijo por defecto si la iProperty de Proyecto está vacía
            End If

            ' 6. Recopilar componentes a revisar y mostrar la ventana interactiva
            Dim items As List(Of ComponenteRenombrarItem) = RenombrarComponentesDialog.RecopilarItems(invApp, oAsmDoc, sPrefijo)
            If items Is Nothing OrElse items.Count = 0 Then
                MsgBox("No se encontraron piezas o ensamblajes para renombrar en este modelo.", MsgBoxStyle.Information, "Renombrar Componentes")
                Exit Sub
            End If

            Using dlg As New RenombrarComponentesDialog(invApp, items, sPrefijo)
                Dim hwnd As IntPtr = IntPtr.Zero
                Try
                    hwnd = New IntPtr(invApp.MainFrameHWND)
                Catch
                End Try

                Dim result As DialogResult
                If hwnd <> IntPtr.Zero Then
                    result = dlg.ShowDialog(New WindowWrapper(hwnd))
                Else
                    result = dlg.ShowDialog()
                End If

                ' 7. Si el usuario cancela, no realizar ningún cambio
                If result <> DialogResult.OK Then
                    Exit Sub
                End If

                ' 8. El usuario presionó Aplicar: ejecutar con los nombres personalizados y sugeridos
                Dim customNames As Dictionary(Of String, String) = dlg.ObtenerNombresPersonalizados()
                Dim renamer As New RenombradorComponentes(invApp)
                renamer.Ejecutar(sRutaActiva, sPrefijo, False, customNames)
            End Using

        Catch ex As Exception
            MsgBox("Error en Renombrar Componentes: " & ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    ' ============================================
    ' OPCIONES DEL ADDIN
    ' ============================================
    Public Shared Sub EjecutarOpciones(
        invApp As Inventor.Application)

        Try
            Using dlg As New OptionsDialog()

                ' Centrar sobre la ventana principal de Inventor
                Dim hwnd As IntPtr = IntPtr.Zero
                Try
                    hwnd = New IntPtr(invApp.MainFrameHWND)
                Catch
                End Try

                If hwnd <> IntPtr.Zero Then
                    dlg.ShowDialog(New WindowWrapper(hwnd))
                Else
                    dlg.ShowDialog()
                End If

            End Using

        Catch ex As Exception
            MsgBox("Error al abrir Opciones: " & ex.Message,
                   MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

End Class