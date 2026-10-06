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
            Return doc.PropertySets.Item("Inventor User Defined Properties").Item(propName).Value.ToString()
        Catch
            Try
                Return doc.PropertySets.Item("Design Tracking Properties").Item(propName).Value.ToString()
            Catch
                Return ""
            End Try
        End Try
    End Function

    Private Shared Sub SetPropertyInDoc(doc As Document, propName As String, value As String)
        Try
            Dim props As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
            Dim prop As Inventor.Property = Nothing
            Try
                prop = props.Item(propName)
            Catch
                ' La propiedad no existe, crearla
                prop = props.Add(value, propName)
            End Try
            If prop IsNot Nothing Then
                prop.Value = value
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

            ' 6. Instanciar la clase y correrla pasándole los datos dinámicos del documento activo
            Dim renamer As New RenombradorComponentes(invApp)

            ' Parámetros: Ruta activa, Prefijo detectado, dryRun (False para que ejecute los cambios reales)
            renamer.Ejecutar(sRutaActiva, sPrefijo, False)

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