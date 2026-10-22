Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

''' <summary>
''' Motor de generación de carpetas y documentación de proyecto (IAM).
''' Traduce y mejora la regla 'Generar_Carpeta.iLogicVb' para el Add-In RuleManager V2.
''' Lee las propiedades OT (Project), Cliente (Vendor), Equipo (Part Number) y Modelo (Stock Number)
''' del ensamblaje activo, crea la carpeta 'CARPETA [OT]' en la ruta del proyecto y genera
''' los archivos de Word (Cartel, Carátula y Lomo) a partir de las plantillas .dotx.
''' </summary>
Public Class GenerarCarpetaEngine

    Private ReadOnly _invApp As Inventor.Application

    Public Sub New(invApp As Inventor.Application)
        _invApp = invApp
    End Sub

    ''' <summary>
    ''' Ejecuta el flujo completo de validación, creación de carpeta y generación de documentos Word.
    ''' </summary>
    Public Sub Ejecutar()
        Try
            ' 1. Validar entorno de ensamblaje activo
            If _invApp.ActiveDocument Is Nothing Then
                MessageBox.Show("No hay ningún documento activo en Autodesk Inventor.",
                                "Generar Carpeta", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If _invApp.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Esta herramienta debe ejecutarse desde un entorno de ensamblaje activo (.iam).",
                                "Generar Carpeta", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim asmDoc As AssemblyDocument = DirectCast(_invApp.ActiveDocument, AssemblyDocument)

            ' 2. Validar que el archivo esté guardado en disco
            If String.IsNullOrWhiteSpace(asmDoc.FullFileName) OrElse Not System.IO.File.Exists(asmDoc.FullFileName) Then
                MessageBox.Show("El ensamblaje activo debe estar guardado en disco antes de generar las carpetas del proyecto.",
                                "Generar Carpeta - Guardar Requerido", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' 3. Extraer propiedades estándar del ensamblaje
            Dim designProps As PropertySet = asmDoc.PropertySets.Item("Design Tracking Properties")

            Dim ot As String = ObtenerPropiedad(designProps, "Project").Trim()
            Dim cliente As String = ObtenerPropiedad(designProps, "Vendor").Trim()
            Dim equipo As String = ObtenerPropiedad(designProps, "Part Number").Trim()
            Dim stockNum As String = ObtenerPropiedad(designProps, "Stock Number").Trim()

            ' 4. Validar propiedades obligatorias (OT y Cliente)
            If String.IsNullOrEmpty(ot) Then
                MessageBox.Show("No se ha asignado OT en Propiedades (Project), inserte la OT antes de continuar.",
                                "Generar Carpeta - OT Requerida", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If

            If String.IsNullOrEmpty(cliente) Then
                MessageBox.Show("No se ha asignado Cliente en Propiedades (Vendor), inserte el Cliente antes de continuar.",
                                "Generar Carpeta - Cliente Requerido", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If

            ' 5. Extraer Modelo (primeros 7 caracteres de Stock Number o completo si tiene menos)
            Dim modelo As String = ""
            If Not String.IsNullOrEmpty(stockNum) Then
                If stockNum.Length >= 7 Then
                    modelo = stockNum.Substring(0, 7)
                Else
                    modelo = stockNum
                End If
            End If

            ' 6. Resolver rutas de proyecto y carpeta de salida
            Dim rutaProyecto As String = System.IO.Path.GetDirectoryName(asmDoc.FullFileName)
            Dim nombreCarpeta As String = "CARPETA " & ot
            Dim rutaCarpeta As String = System.IO.Path.Combine(rutaProyecto, nombreCarpeta)

            If Not System.IO.Directory.Exists(rutaCarpeta) Then
                System.IO.Directory.CreateDirectory(rutaCarpeta)
            End If

            ' 7. Nombres de archivos de destino
            Dim nombreArchivoCartel As String = ot & " Cartel.docx"
            Dim nombreArchivoCaratula As String = ot & " Caratula.docx"
            Dim nombreArchivoLomo As String = ot & " Lomo.docx"

            Dim rutaSalidaCartel As String = System.IO.Path.Combine(rutaCarpeta, nombreArchivoCartel)
            Dim rutaSalidaCaratula As String = System.IO.Path.Combine(rutaCarpeta, nombreArchivoCaratula)
            Dim rutaSalidaLomo As String = System.IO.Path.Combine(rutaCarpeta, nombreArchivoLomo)

            ' 8. Resolver rutas de las plantillas .dotx
            Dim rutaPlantillaCartel As String = ResolverRutaPlantilla("CARTEL_PROYECTO.dotx", rutaProyecto)
            Dim rutaPlantillaCaratula As String = ResolverRutaPlantilla("CARATULAS_CARPETAS.dotx", rutaProyecto)
            Dim rutaPlantillaLomo As String = ResolverRutaPlantilla("LOMO_CARPETAS.dotx", rutaProyecto)

            If String.IsNullOrEmpty(rutaPlantillaCartel) OrElse
               String.IsNullOrEmpty(rutaPlantillaCaratula) OrElse
               String.IsNullOrEmpty(rutaPlantillaLomo) Then
                ' Mensaje ya emitido en ResolverRutaPlantilla si se canceló
                Return
            End If

            ' 9. Generar documentos con Microsoft Word (automatización COM protegida)
            GenerarDocumentosWord(
                ot, cliente, equipo, modelo,
                rutaPlantillaCartel, rutaSalidaCartel,
                rutaPlantillaCaratula, rutaSalidaCaratula,
                rutaPlantillaLomo, rutaSalidaLomo
            )

            ' 10. Notificación y apertura opcional de la carpeta generada
            Dim mensajeExito As String =
                "Se han generado con éxito los elementos de la carpeta de proyecto:" & vbCrLf & vbCrLf &
                "• Carpeta: " & rutaCarpeta & vbCrLf &
                "• " & nombreArchivoCartel & vbCrLf &
                "• " & nombreArchivoCaratula & vbCrLf &
                "• " & nombreArchivoLomo & vbCrLf & vbCrLf &
                "¿Desea abrir la carpeta en el Explorador de Windows ahora?"

            Dim respuesta As DialogResult = MessageBox.Show(
                mensajeExito,
                "Generar Carpeta - Proceso Terminado",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information
            )

            If respuesta = DialogResult.Yes Then
                Try
                    Process.Start(New ProcessStartInfo("explorer.exe", """" & rutaCarpeta & """") With {
                        .UseShellExecute = True
                    })
                Catch exExplorer As Exception
                    ' Ignorar si no se puede abrir el explorador
                End Try
            End If

        Catch ex As Exception
            MessageBox.Show("Error durante la generación de la carpeta: " & vbCrLf & ex.Message,
                            "Generar Carpeta - Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Automatiza Microsoft Word para crear y rellenar los marcadores de los 3 documentos.
    ''' </summary>
    Private Sub GenerarDocumentosWord(
        ot As String,
        cliente As String,
        equipo As String,
        modelo As String,
        rutaPlantillaCartel As String,
        rutaSalidaCartel As String,
        rutaPlantillaCaratula As String,
        rutaSalidaCaratula As String,
        rutaPlantillaLomo As String,
        rutaSalidaLomo As String)

        Dim wordApp As Object = Nothing
        Dim docActual As Object = Nothing

        Try
            ' Crear instancia de Word mediante Late-Binding COM
            Try
                wordApp = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"))
            Catch exWordInst As Exception
                wordApp = Interaction.CreateObject("Word.Application")
            End Try

            If wordApp Is Nothing Then
                Throw New Exception("No se pudo iniciar Microsoft Word. Verifique que Microsoft Office/Word esté instalado.")
            End If

            wordApp.Visible = False
            wordApp.DisplayAlerts = 0 ' wdAlertsNone

            ' Diccionario de marcadores estándar
            Dim marcadoresCompletos As New Dictionary(Of String, String) From {
                {"MODELO", modelo},
                {"EQUIPO", equipo},
                {"OT", ot},
                {"CLIENTE", cliente}
            }

            ' 1. Procesar CARTEL
            docActual = wordApp.Documents.Add(rutaPlantillaCartel)
            RellenarMarcadores(docActual, marcadoresCompletos)
            docActual.SaveAs2(rutaSalidaCartel)
            docActual.Close()
            LiberarComObject(docActual)
            docActual = Nothing

            ' 2. Procesar CARÁTULA
            docActual = wordApp.Documents.Add(rutaPlantillaCaratula)
            RellenarMarcadores(docActual, marcadoresCompletos)
            docActual.SaveAs2(rutaSalidaCaratula)
            docActual.Close()
            LiberarComObject(docActual)
            docActual = Nothing

            ' 3. Procesar LOMO
            docActual = wordApp.Documents.Add(rutaPlantillaLomo)
            RellenarMarcadores(docActual, marcadoresCompletos)
            docActual.SaveAs2(rutaSalidaLomo)
            docActual.Close()
            LiberarComObject(docActual)
            docActual = Nothing

        Finally
            ' Cierre garantizado de la aplicación de Word y recolección de COM
            If docActual IsNot Nothing Then
                Try
                    docActual.Close(False)
                Catch
                End Try
                LiberarComObject(docActual)
                docActual = Nothing
            End If

            If wordApp IsNot Nothing Then
                Try
                    wordApp.Quit()
                Catch
                End Try
                LiberarComObject(wordApp)
                wordApp = Nothing
            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    ''' <summary>
    ''' Rellena los marcadores existentes en el documento Word de forma segura.
    ''' </summary>
    Private Sub RellenarMarcadores(docWord As Object, marcadores As Dictionary(Of String, String))
        If docWord Is Nothing OrElse marcadores Is Nothing Then Return

        For Each kvp In marcadores
            Try
                If CBool(docWord.Bookmarks.Exists(kvp.Key)) Then
                    Dim bm As Object = docWord.Bookmarks.Item(kvp.Key)
                    Dim rng As Object = bm.Range
                    rng.Text = If(kvp.Value, "")
                End If
            Catch ex As Exception
                ' Si un marcador específico falla, continuar con los demás
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Resuelve la ruta de una plantilla .dotx buscando en la configuración, ruta original y proyecto.
    ''' Si no se encuentra, permite al usuario seleccionarla manualmente.
    ''' </summary>
    Private Function ResolverRutaPlantilla(nombreArchivo As String, rutaProyecto As String) As String
        ' 1. Buscar en configuración del Add-in
        Dim rutaConfig As String = AddinSettings.Current.RutaPlantillasReglas
        If Not String.IsNullOrEmpty(rutaConfig) Then
            Dim c1 As String = System.IO.Path.Combine(rutaConfig, nombreArchivo)
            If System.IO.File.Exists(c1) Then Return c1
        End If

        ' 2. Ruta original por defecto (D:\Bibliotecas\Ilogic\Plantillas para reglas)
        Dim c2 As String = System.IO.Path.Combine("D:\Bibliotecas\Ilogic\Plantillas para reglas", nombreArchivo)
        If System.IO.File.Exists(c2) Then Return c2

        ' 3. Carpeta 'Plantillas' dentro de la ruta del proyecto o add-in
        Dim c3 As String = System.IO.Path.Combine(rutaProyecto, "Plantillas", nombreArchivo)
        If System.IO.File.Exists(c3) Then Return c3

        Dim addinDir As String = System.IO.Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly().Location)
        Dim c4 As String = System.IO.Path.Combine(addinDir, "Plantillas", nombreArchivo)
        If System.IO.File.Exists(c4) Then Return c4

        ' 4. Si no se encuentra automáticamente, solicitar al usuario que ubique el archivo
        Dim prompt As String = "No se encontró la plantilla requerida:" & vbCrLf &
                               nombreArchivo & vbCrLf & vbCrLf &
                               "¿Desea localizarla manualmente en el explorador de archivos?"

        If MessageBox.Show(prompt, "Plantilla No Encontrada", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Using ofd As New OpenFileDialog()
                ofd.Title = "Seleccione la plantilla " & nombreArchivo
                ofd.Filter = "Plantillas de Word (*.dotx;*.dot;*.docx)|*.dotx;*.dot;*.docx|Todos los archivos (*.*)|*.*"
                ofd.FileName = nombreArchivo
                If ofd.ShowDialog() = DialogResult.OK AndAlso System.IO.File.Exists(ofd.FileName) Then
                    ' Guardar la carpeta elegida para agilizar futuras búsquedas
                    Dim nuevaCarpeta As String = System.IO.Path.GetDirectoryName(ofd.FileName)
                    AddinSettings.Current.RutaPlantillasReglas = nuevaCarpeta
                    AddinSettings.Save()
                    Return ofd.FileName
                End If
            End Using
        End If

        Return ""
    End Function

    ''' <summary>
    ''' Obtiene el valor textual de una propiedad de Inventor de forma segura.
    ''' </summary>
    Private Function ObtenerPropiedad(propSet As PropertySet, nombreProp As String) As String
        Try
            If propSet IsNot Nothing Then
                Dim prop As [Property] = propSet.Item(nombreProp)
                If prop IsNot Nothing AndAlso prop.Value IsNot Nothing Then
                    Return prop.Value.ToString()
                End If
            End If
        Catch
        End Try
        Return ""
    End Function

    ''' <summary>
    ''' Libera un objeto COM de memoria de forma limpia.
    ''' </summary>
    Private Sub LiberarComObject(ByRef obj As Object)
        If obj IsNot Nothing Then
            Try
                Marshal.ReleaseComObject(obj)
            Catch
            Finally
                obj = Nothing
            End Try
        End If
    End Sub

End Class
